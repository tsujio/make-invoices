package main

import (
	"context"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/calendar/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

func getPathSiblingOfExecutable(filename string) string {
	exe, err := os.Executable()
	if err != nil {
		log.Fatalf("Failed to get executable path: %v", err)
	}
	return filepath.Join(filepath.Dir(exe), filename)
}

type Config struct {
	CredentialsFileName    string   `json:"credentials_file_name"`
	OAuth2TokenFileName    string   `json:"oauth2_token_file_name"`
	CalendarID             string   `json:"calendar_id"`
	WorkDayTitle           string   `json:"work_day_title"`
	WorkStartTime          string   `json:"work_start_time"`
	WorkSpreadsheetIDs     []string `json:"work_spreadsheet_ids"`
	WorkDocumentTemplateID string   `json:"work_document_template_id"`
}

func loadConfig() *Config {
	path := getPathSiblingOfExecutable("config.json")
	f, err := os.Open(path)
	if err != nil {
		log.Fatalf("Failed to open config file: %v", err)
	}
	var config Config
	if err := json.NewDecoder(f).Decode(&config); err != nil {
		log.Fatalf("Failed to decode config file: %v", err)
	}
	return &config
}

func createAPIClient(ctx context.Context, config *Config) *http.Client {
	// Create OAuth2 config
	cred, err := ioutil.ReadFile(getPathSiblingOfExecutable(config.CredentialsFileName))
	if err != nil {
		log.Fatalf("Failed to read credentials file: %v", err)
	}
	oauth2Conf, err := google.ConfigFromJSON(
		cred,
		calendar.CalendarReadonlyScope,
		"https://www.googleapis.com/auth/spreadsheets",
		"https://www.googleapis.com/auth/drive",
	)
	if err != nil {
		log.Fatalf("Failed to make oauth2 config from json: %v", err)
	}

	// Get oauth token
	var token *oauth2.Token
	tokenFilePath := getPathSiblingOfExecutable(config.OAuth2TokenFileName)
	if f, err := os.Open(tokenFilePath); err == nil {
		// From local file (if exists)
		defer f.Close()
		tok := oauth2.Token{}
		err := json.NewDecoder(f).Decode(&tok)
		if err != nil {
			log.Fatalf("Failed to decode oauth token: %v", err)
		}
		token = &tok
	} else {
		// From web
		authURL := oauth2Conf.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
		fmt.Printf("Go to the following link in your browser then type the authorization code: \n%v\n", authURL)
		var authCode string
		fmt.Printf("Code: ")
		if _, err := fmt.Scan(&authCode); err != nil {
			log.Fatalf("Unable to read authorization code: %v", err)
		}
		tok, err := oauth2Conf.Exchange(context.TODO(), authCode)
		if err != nil {
			log.Fatalf("Unable to retrieve token from web: %v", err)
		}
		f, err := os.OpenFile(tokenFilePath, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
		if err != nil {
			log.Fatalf("Unable to cache oauth token: %v", err)
		}
		defer f.Close()
		json.NewEncoder(f).Encode(tok)
		token = tok
	}

	return oauth2Conf.Client(ctx, token)
}

func getCalendarSchedules(ctx context.Context, client *http.Client, calendarID string, targetTime time.Time, filter func(*calendar.Event) bool) []time.Time {
	cal, err := calendar.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Failed to create calendar client: %v", err)
	}

	// Fetch calendar items
	events, err := cal.Events.List(calendarID).
		ShowDeleted(false).
		SingleEvents(true).
		TimeMin(targetTime.AddDate(0, -1, -1).Format(time.RFC3339)).
		MaxResults(999).
		OrderBy("startTime").
		Do()
	if err != nil {
		log.Fatalf("Failed to retrieve calendar items: %v", err)
	}

	// Collect items
	items := make([]time.Time, 0)
	for _, item := range events.Items {
		var date time.Time
		var err error
		if item.Start.DateTime != "" {
			date, err = time.Parse(time.RFC3339, item.Start.DateTime)
			if err != nil {
				log.Fatalf("Failed to parse calendar datetime: %v", err)
			}
		} else {
			date, err = time.Parse("2006-01-02", item.Start.Date)
			if err != nil {
				log.Fatalf("Failed to parse calendar date: %v", err)
			}
		}
		if date.Year() != targetTime.Year() || date.Month() != targetTime.Month() {
			continue
		}

		if filter != nil && !filter(item) {
			continue
		}

		items = append(items, date)
	}

	return items
}

func updateAndDownloadWorkSpreadsheets(ctx context.Context, client *http.Client, targetTime time.Time, workDays []time.Time, config *Config) {
	sht, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Failed to create sheet client: %v", err)
	}

	for _, spreadsheetID := range config.WorkSpreadsheetIDs {
		// Get spreadsheet
		spreadsheet, err := sht.Spreadsheets.Get(spreadsheetID).Do()
		if err != nil {
			log.Fatalf("Failed to get spreadsheet: %v", err)
		}

		// Get sheet for targetTime
		var targetSheetID int64
		for _, s := range spreadsheet.Sheets {
			if targetTime.Format("200601") == s.Properties.Title {
				// Already exists
				targetSheetID = s.Properties.SheetId
			}
		}
		if targetSheetID == 0 {
			// Copy from latest sheet if target sheet not found
			var copyFrom *sheets.Sheet
			for _, s := range spreadsheet.Sheets {
				if targetTime.AddDate(0, -1, 0).Format("200601") == s.Properties.Title {
					copyFrom = s
					break
				}
			}
			if copyFrom == nil {
				log.Fatalf("Failed to determine sheet to copy")
			}
			dest, err := sht.Spreadsheets.Sheets.CopyTo(spreadsheetID, copyFrom.Properties.SheetId, &sheets.CopySheetToAnotherSpreadsheetRequest{
				DestinationSpreadsheetId: spreadsheetID,
			}).Do()
			if err != nil {
				log.Fatalf("Failed to copy sheet: %v", err)
			}
			targetSheetID = dest.SheetId
			if _, err := sht.Spreadsheets.BatchUpdate(spreadsheetID, &sheets.BatchUpdateSpreadsheetRequest{
				Requests: []*sheets.Request{{
					UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
						Fields: "title,index",
						Properties: &sheets.SheetProperties{
							SheetId: targetSheetID,
							Title:   targetTime.Format("200601"),
							Index:   0,
						},
					},
				}},
			}).Do(); err != nil {
				log.Fatalf("Failed to update sheet position: %v", err)
			}
		}

		// Update date
		if _, err := sht.Spreadsheets.Values.Update(spreadsheetID, targetTime.Format("200601")+"!M3:M3", &sheets.ValueRange{
			Values: [][]interface{}{{targetTime.Format("2006/01/02")}},
		}).ValueInputOption("USER_ENTERED").Do(); err != nil {
			log.Fatalf("Failed to set work month to sheet: %v", err)
		}

		// Update work times
		values := make([][]interface{}, 0)
		for i := 1; i <= 31; i++ {
			value := ""
			for _, d := range workDays {
				if targetTime.Year() == d.Year() && targetTime.Month() == d.Month() && i == d.Day() {
					value = config.WorkStartTime
					break
				}
			}
			values = append(values, []interface{}{value})
		}
		if _, err := sht.Spreadsheets.Values.Update(spreadsheetID, targetTime.Format("200601")+"!D7:D37", &sheets.ValueRange{
			Values: values,
		}).ValueInputOption("USER_ENTERED").Do(); err != nil {
			log.Fatalf("Failed to set work times to sheet: %v", err)
		}

		// Export to pdf
		resp, err := client.Get(fmt.Sprintf("https://docs.google.com/spreadsheets/d/%s/export?format=pdf&gid=%d", spreadsheetID, targetSheetID))
		if err != nil {
			log.Fatalf("Failed to export spreadsheet: %v", err)
		}
		defer resp.Body.Close()
		d, err := ioutil.ReadAll(resp.Body)
		if err != nil {
			log.Fatalf("Failed to read response for export spreadsheet: %v", err)
		}
		if err := ioutil.WriteFile(fmt.Sprintf("%s%s.pdf", targetTime.Format("200601"), spreadsheet.Properties.Title), d, 0666); err != nil {
			log.Fatalf("Failed to save spreadsheet pdf: %v", err)
		}
	}
}

func main() {
	jst, err := time.LoadLocation("Asia/Tokyo")
	if err != nil {
		log.Fatalf("Failed to load timezone: %v", err)
	}
	var targetTime time.Time
	if len(os.Args) < 2 {
		targetTime = time.Now().In(jst)
	} else {
		var err error
		targetTime, err = time.Parse("200601", os.Args[1])
		if err != nil {
			log.Fatalf("Failed to parse date parameter: %v", err)
		}
	}
	targetTime = time.Date(targetTime.Year(), targetTime.Month(), 1, 0, 0, 0, 0, jst)

	log.Printf("Make invoices for %s? (Y/n): ", targetTime.Format("200601"))
	var ans string
	fmt.Scanln(&ans)
	if ans = strings.TrimSuffix(ans, "\n"); ans != "" && strings.ToLower(ans) != "y" {
		os.Exit(0)
	}

	ctx := context.Background()

	config := loadConfig()

	log.Println("Loaded config")

	client := createAPIClient(ctx, config)

	workDays := getCalendarSchedules(ctx, client, config.CalendarID, targetTime, func(e *calendar.Event) bool {
		return e.Summary == config.WorkDayTitle
	})

	log.Printf("Found %d work days\n", len(workDays))

	updateAndDownloadWorkSpreadsheets(ctx, client, targetTime, workDays, config)

	log.Println("Exported spreadsheets")

	log.Println("Done")
}
