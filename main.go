package main

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"net/http"
	"os"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type TokenResponse struct {
	AccessToken  string `json:"access_token"`
	RefreshToken string `json:"refresh_token"`
	ExpiresAt    int64  `json:"expires_at"`
}

type StravaActivity struct {
	ID             int64   `json:"id"`
	Name           string  `json:"name"`
	Type           string  `json:"type"`        // legacy
	SportType      string  `json:"sport_type"`  // newer
	Distance       float64 `json:"distance"`    // meters
	MovingTime     int     `json:"moving_time"` // seconds
	ElapsedTime    int     `json:"elapsed_time"`
	AverageSpeed   float64 `json:"average_speed"` // m/s (may be 0 sometimes)
	StartDate      string  `json:"start_date"`    // UTC ISO
	StartDateLocal string  `json:"start_date_local"`
	Timezone       string  `json:"timezone"`
}

type DayStats struct {
	TotalMeters  int
	TotalSeconds int
}

func mustEnv(k string) string {
	v := strings.TrimSpace(os.Getenv(k))
	if v == "" {
		panic("missing env var: " + k)
	}
	return v
}

// mustCell converts col/row coords to an Excel cell (e.g., 1,1 -> "A1").
// Panics on impossible coordinates (fine for a CLI exporter).
func mustCell(col, row int) string {
	c, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		panic(err)
	}
	return c
}

func refreshAccessToken(clientID, clientSecret, refreshToken string) (TokenResponse, error) {
	form := map[string]string{
		"client_id":     clientID,
		"client_secret": clientSecret,
		"grant_type":    "refresh_token",
		"refresh_token": refreshToken,
	}
	body, _ := json.Marshal(form)

	req, _ := http.NewRequest("POST", "https://www.strava.com/oauth/token", bytes.NewReader(body))
	req.Header.Set("Content-Type", "application/json")

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return TokenResponse{}, err
	}
	defer resp.Body.Close()

	if resp.StatusCode != 200 {
		return TokenResponse{}, fmt.Errorf("token refresh failed: http %d", resp.StatusCode)
	}

	var tr TokenResponse
	if err := json.NewDecoder(resp.Body).Decode(&tr); err != nil {
		return TokenResponse{}, err
	}
	return tr, nil
}

// Strava: GET /api/v3/athlete/activities?per_page=200&page=N
func listActivities(accessToken string) ([]StravaActivity, error) {
	var out []StravaActivity
	page := 1
	perPage := 200

	throttle := time.NewTicker(250 * time.Millisecond)
	defer throttle.Stop()

	for {
		<-throttle.C
		url := fmt.Sprintf("https://www.strava.com/api/v3/athlete/activities?per_page=%d&page=%d", perPage, page)
		req, _ := http.NewRequest("GET", url, nil)
		req.Header.Set("Authorization", "Bearer "+accessToken)

		resp, err := http.DefaultClient.Do(req)
		if err != nil {
			return nil, err
		}

		if resp.StatusCode == 429 {
			resp.Body.Close()
			time.Sleep(20 * time.Second)
			continue
		}
		if resp.StatusCode != 200 {
			resp.Body.Close()
			return nil, fmt.Errorf("activities request failed: http %d", resp.StatusCode)
		}

		var batch []StravaActivity
		if err := json.NewDecoder(resp.Body).Decode(&batch); err != nil {
			resp.Body.Close()
			return nil, err
		}
		resp.Body.Close()

		if len(batch) == 0 {
			break
		}
		out = append(out, batch...)
		page++
	}
	return out, nil
}

func isRun(a StravaActivity) bool {
	t := strings.ToLower(strings.TrimSpace(a.SportType))
	if t == "" {
		t = strings.ToLower(strings.TrimSpace(a.Type))
	}
	switch t {
	case "run", "trailrun", "virtualrun":
		return true
	default:
		return false
	}
}

func parseLocalDate(a StravaActivity) (time.Time, error) {
	s := strings.TrimSpace(a.StartDateLocal)
	if s == "" {
		s = strings.TrimSpace(a.StartDate)
	}
	if s == "" {
		return time.Time{}, errors.New("missing start_date_local/start_date")
	}

	// Try RFC3339 first
	if t, err := time.Parse(time.RFC3339, s); err == nil {
		return t, nil
	}
	// Try without timezone
	if t, err := time.Parse("2006-01-02T15:04:05", s); err == nil {
		return t, nil
	}
	return time.Time{}, fmt.Errorf("cannot parse date: %q", s)
}

func paceString(totalSeconds int, totalMeters int) string {
	if totalMeters <= 0 || totalSeconds <= 0 {
		return ""
	}
	km := float64(totalMeters) / 1000.0
	secPerKm := float64(totalSeconds) / km
	min := int(secPerKm) / 60
	sec := int(secPerKm) % 60
	return fmt.Sprintf("%d:%02d", min, sec)
}

func kmString(totalMeters int) string {
	return fmt.Sprintf("%.2f", float64(totalMeters)/1000.0)
}

func buildDailyStats(activities []StravaActivity) map[string]DayStats {
	// key: YYYY-MM-DD
	m := map[string]DayStats{}
	for _, a := range activities {
		if !isRun(a) {
			continue
		}
		dt, err := parseLocalDate(a)
		if err != nil {
			continue
		}
		key := dt.Format("2006-01-02")
		ds := m[key]
		ds.TotalMeters += int(a.Distance)
		ds.TotalSeconds += a.MovingTime
		m[key] = ds
	}
	return m
}

func ensureSheet(f *excelize.File, name string) error {
	idx, err := f.GetSheetIndex(name) // (int, error)
	if err == nil && idx != -1 {
		return nil
	}
	f.NewSheet(name)
	return nil
}

func setColWidths(f *excelize.File, sheet string, startCol, endCol string, w float64) {
	_ = f.SetColWidth(sheet, startCol, endCol, w)
}

func calendarLayoutForYear(f *excelize.File, sheet string, year int, daily map[string]DayStats) error {
	monthsPerRow := 3
	blockW := 8 // columns per month block (we use 7)
	blockH := 9 // rows per month block: title + weekday header + 6 weeks

	startRow := 1
	startCol := 1

	titleStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 14},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	monthTitleStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 12},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	weekdayStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 10},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "bottom", Style: 1},
		},
	})
	dayCellStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Size: 9},
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "top",
			WrapText:   true,
		},
		Border: []excelize.Border{
			{Type: "left", Style: 1},
			{Type: "right", Style: 1},
			{Type: "top", Style: 1},
			{Type: "bottom", Style: 1},
		},
	})

	// Big year title across the whole used area
	yearTitleCell := mustCell(1, 1)
	lastTitleCell := mustCell(monthsPerRow*blockW, 1)
	_ = f.MergeCell(sheet, yearTitleCell, lastTitleCell)
	_ = f.SetCellValue(sheet, yearTitleCell, fmt.Sprintf("%d Runs Calendar", year))
	_ = f.SetCellStyle(sheet, yearTitleCell, lastTitleCell, titleStyle)

	weekdays := []string{"Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"}

	// Basic widths/heights
	setColWidths(f, sheet, "A", "XFD", 2.0) // broad reset
	for c := 1; c <= monthsPerRow*blockW; c++ {
		colName, _ := excelize.ColumnNumberToName(c)
		_ = f.SetColWidth(sheet, colName, colName, 16)
	}
	for r := 1; r <= startRow+4*blockH+2; r++ {
		_ = f.SetRowHeight(sheet, r, 42)
	}
	_ = f.SetRowHeight(sheet, 1, 26)

	for month := 1; month <= 12; month++ {
		monthIndex := month - 1
		gridRow := monthIndex / monthsPerRow
		gridCol := monthIndex % monthsPerRow

		topRow := startRow + 1 + gridRow*blockH
		leftCol := startCol + gridCol*blockW

		// Month title merged across 7 weekday columns
		monthTitleStart := mustCell(leftCol, topRow)
		monthTitleEnd := mustCell(leftCol+6, topRow)
		_ = f.MergeCell(sheet, monthTitleStart, monthTitleEnd)

		monthName := time.Date(year, time.Month(month), 1, 0, 0, 0, 0, time.Local).Format("January")
		_ = f.SetCellValue(sheet, monthTitleStart, monthName)
		_ = f.SetCellStyle(sheet, monthTitleStart, monthTitleEnd, monthTitleStyle)
		_ = f.SetRowHeight(sheet, topRow, 22)

		// Weekday header
		for i := 0; i < 7; i++ {
			cell := mustCell(leftCol+i, topRow+1)
			_ = f.SetCellValue(sheet, cell, weekdays[i])
			_ = f.SetCellStyle(sheet, cell, cell, weekdayStyle)
			_ = f.SetRowHeight(sheet, topRow+1, 18)
		}

		// Days grid (6 weeks)
		first := time.Date(year, time.Month(month), 1, 0, 0, 0, 0, time.Local)

		// Go: Sunday=0..Saturday=6 -> Monday-first: Monday=0..Sunday=6
		offset := (int(first.Weekday()) + 6) % 7

		daysInMonth := time.Date(year, time.Month(month)+1, 0, 0, 0, 0, 0, time.Local).Day()

		day := 1
		for week := 0; week < 6; week++ {
			for dow := 0; dow < 7; dow++ {
				r := topRow + 2 + week
				c := leftCol + dow
				cell := mustCell(c, r)

				gridIndex := week*7 + dow
				if gridIndex < offset || day > daysInMonth {
					_ = f.SetCellValue(sheet, cell, "")
					_ = f.SetCellStyle(sheet, cell, cell, dayCellStyle)
					continue
				}

				date := time.Date(year, time.Month(month), day, 0, 0, 0, 0, time.Local)
				key := date.Format("2006-01-02")
				ds, ok := daily[key]

				if ok && ds.TotalMeters > 0 && ds.TotalSeconds > 0 {
					text := fmt.Sprintf("%d\n%s km\n%s /km",
						day,
						kmString(ds.TotalMeters),
						paceString(ds.TotalSeconds, ds.TotalMeters),
					)
					_ = f.SetCellValue(sheet, cell, text)
				} else {
					_ = f.SetCellValue(sheet, cell, fmt.Sprintf("%d", day))
				}

				_ = f.SetCellStyle(sheet, cell, cell, dayCellStyle)
				day++
			}
		}
	}

	return nil
}

func main() {
	clientID := mustEnv("STRAVA_CLIENT_ID")
	clientSecret := mustEnv("STRAVA_CLIENT_SECRET")
	refreshToken := mustEnv("STRAVA_REFRESH_TOKEN")

	outFile := strings.TrimSpace(os.Getenv("OUT_XLSX"))
	if outFile == "" {
		outFile = "strava_runs_calendar.xlsx"
	}

	tr, err := refreshAccessToken(clientID, clientSecret, refreshToken)
	if err != nil {
		panic(err)
	}

	fmt.Println("Access token expires at:", time.Unix(tr.ExpiresAt, 0).Format(time.RFC3339))
	if tr.RefreshToken != "" && tr.RefreshToken != refreshToken {
		fmt.Println("NOTE: refresh token rotated. New STRAVA_REFRESH_TOKEN =", tr.RefreshToken)
	}

	acts, err := listActivities(tr.AccessToken)
	if err != nil {
		panic(err)
	}
	fmt.Println("Fetched activities:", len(acts))

	var runs []StravaActivity
	for _, a := range acts {
		if isRun(a) {
			runs = append(runs, a)
		}
	}
	sort.Slice(runs, func(i, j int) bool { return runs[i].StartDateLocal < runs[j].StartDateLocal })
	fmt.Println("Run-like activities:", len(runs))

	daily := buildDailyStats(runs)

	// Determine year range from runs
	minYear := 9999
	maxYear := 0
	for k := range daily {
		t, err := time.Parse("2006-01-02", k)
		if err != nil {
			continue
		}
		if t.Year() < minYear {
			minYear = t.Year()
		}
		if t.Year() > maxYear {
			maxYear = t.Year()
		}
	}
	if maxYear == 0 {
		panic("no run data found")
	}

	f := excelize.NewFile()

	firstSheet := strconv.Itoa(minYear)

	// Default sheet is usually "Sheet1". Rename it to the first year (and handle error).
	if err := f.SetSheetName("Sheet1", firstSheet); err != nil {
		// If rename fails for any reason, just create the first year sheet.
		f.NewSheet(firstSheet)
	}

	for y := minYear; y <= maxYear; y++ {
		sheet := strconv.Itoa(y)
		if err := ensureSheet(f, sheet); err != nil {
			panic(err)
		}
		if err := calendarLayoutForYear(f, sheet, y, daily); err != nil {
			panic(err)
		}
	}

	// Set first year as active (GetSheetIndex returns (int, error))
	idx, err := f.GetSheetIndex(firstSheet)
	if err == nil && idx != -1 {
		f.SetActiveSheet(idx)
	}

	if err := f.SaveAs(outFile); err != nil {
		panic(err)
	}
	fmt.Println("Wrote:", outFile)
}
