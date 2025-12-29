package xlrd

import (
	"testing"
	"time"
)

const (
	not1904 = 0 // 1900-based
	is1904  = 1 // 1904-based
)

func TestDatesAndTimes1900Epoch(t *testing.T) {
	// Test Excel dates strings and corresponding serial date numbers taken
	// from an Excel file.
	excelDates := []struct {
		expected string
		xldate   float64
	}{
		// Excel's 0.0 date in the 1900 epoch is 1 day before 1900.
		{"1899-12-31T00:00:00.000", 0},
		// Date/time before the false Excel 1900 leapday.
		{"1900-02-28T02:11:11.986", 59.09111094906},
		// Date/time after the false Excel 1900 leapday.
		{"1900-03-01T05:46:44.068", 61.24078782403},
		// Random date/times in Excel's 0-9999.9999+ range.
		{"1982-08-25T00:15:20.213", 30188.010650613425},
		{"2065-04-19T00:16:48.290", 60376.011670023145},
		{"3222-06-11T03:08:08.251", 483014.13065105322},
		{"4379-08-03T06:14:48.580", 905652.26028449077},
		{"5949-12-30T12:59:54.263", 1479232.5416002662},
		// End of Excel's date range.
		{"9999-12-31T23:59:59.000", 2958465.999988426},
	}

	for _, tt := range excelDates {
		exp, err := time.Parse("2006-01-02T15:04:05.000", tt.expected)
		if err != nil {
			t.Fatalf("Failed to parse expected time %s: %v", tt.expected, err)
		}
		got, err := XldateAsDatetime(tt.xldate, not1904)
		if err != nil {
			t.Errorf("XldateAsDatetime(%f, %v) error = %v", tt.xldate, not1904, err)
			continue
		}
		if !got.Equal(exp) {
			t.Errorf("XldateAsDatetime(%f, %v) = %v, want %v", tt.xldate, not1904, got, exp)
		}
	}
}

func TestDatesOnly1900Epoch(t *testing.T) {
	// Test Excel dates strings and corresponding serial date numbers taken
	// from an Excel file.
	excelDates := []struct {
		expected string
		xldate   float64
	}{
		// Excel's day 0 in the 1900 epoch is 1 day before 1900.
		{"1899-12-31", 0},
		// Excel's day 1 in the 1900 epoch.
		{"1900-01-01", 1},
		// Date/time before the false Excel 1900 leapday.
		{"1900-02-28", 59},
		// Date/time after the false Excel 1900 leapday.
		{"1900-03-01", 61},
		// Random date/times in Excel's 0-9999.9999+ range.
		{"1902-09-27", 1001},
		{"1999-12-31", 36525},
		{"2000-01-01", 36526},
		{"4000-12-31", 767376},
		{"4321-01-01", 884254},
		{"9999-01-01", 2958101},
		// End of Excel's date range.
		{"9999-12-31", 2958465},
	}

	for _, tt := range excelDates {
		exp, err := time.Parse("2006-01-02", tt.expected)
		if err != nil {
			t.Fatalf("Failed to parse expected time %s: %v", tt.expected, err)
		}
		got, err := XldateAsDatetime(tt.xldate, not1904)
		if err != nil {
			t.Errorf("XldateAsDatetime(%f, %v) error = %v", tt.xldate, not1904, err)
			continue
		}
		// Compare only date part (ignore time)
		expDate := time.Date(exp.Year(), exp.Month(), exp.Day(), 0, 0, 0, 0, time.UTC)
		gotDate := time.Date(got.Year(), got.Month(), got.Day(), 0, 0, 0, 0, time.UTC)
		if !gotDate.Equal(expDate) {
			t.Errorf("XldateAsDatetime(%f, %v) = %v, want %v", tt.xldate, not1904, gotDate, expDate)
		}
	}
}

func TestDatesOnly1904Epoch(t *testing.T) {
	// Test Excel dates strings and corresponding serial date numbers taken
	// from an Excel file.
	excelDates := []struct {
		expected string
		xldate   float64
	}{
		// Excel's day 0 in the 1904 epoch.
		{"1904-01-01", 0},
		// Random date/times in Excel's 0-9999.9999+ range.
		{"1904-01-31", 30},
		{"1904-08-31", 243},
		{"1999-02-28", 34757},
		{"1999-12-31", 35063},
		{"2000-01-01", 35064},
		{"2400-12-31", 181526},
		{"4000-01-01", 765549},
		{"9999-01-01", 2956639},
		// End of Excel's date range.
		{"9999-12-31", 2957003},
	}

	for _, tt := range excelDates {
		exp, err := time.Parse("2006-01-02", tt.expected)
		if err != nil {
			t.Fatalf("Failed to parse expected time %s: %v", tt.expected, err)
		}
		got, err := XldateAsDatetime(tt.xldate, is1904)
		if err != nil {
			t.Errorf("XldateAsDatetime(%f, %v) error = %v", tt.xldate, is1904, err)
			continue
		}
		// Compare only date part (ignore time)
		expDate := time.Date(exp.Year(), exp.Month(), exp.Day(), 0, 0, 0, 0, time.UTC)
		gotDate := time.Date(got.Year(), got.Month(), got.Day(), 0, 0, 0, 0, time.UTC)
		if !gotDate.Equal(expDate) {
			t.Errorf("XldateAsDatetime(%f, %v) = %v, want %v", tt.xldate, is1904, gotDate, expDate)
		}
	}
}

func TestTimesOnly(t *testing.T) {
	// Test Excel dates strings and corresponding serial date numbers taken
	// from an Excel file. The 1899-12-31 date is Excel's day 0.
	excelDates := []struct {
		expected string
		xldate   float64
	}{
		// Random times in Excel's 0-0.9999+ range for 1 day.
		{"1899-12-31T00:00:00.000", 0},
		{"1899-12-31T00:15:20.213", 1.0650613425925924e-2},
		{"1899-12-31T02:24:37.095", 0.10042934027777778},
		{"1899-12-31T04:56:35.792", 0.2059698148148148},
		{"1899-12-31T07:31:20.407", 0.31343063657407405},
		{"1899-12-31T09:37:23.945", 0.40097158564814817},
		{"1899-12-31T12:09:48.602", 0.50681252314814818},
		{"1899-12-31T14:37:57.451", 0.60969271990740748},
		{"1899-12-31T17:04:02.415", 0.71113906250000003},
		{"1899-12-31T19:14:24.673", 0.80167445601851861},
		{"1899-12-31T21:39:05.944", 0.90215212962962965},
		{"1899-12-31T23:17:12.632", 0.97028509259259266},
		{"1899-12-31T23:59:59.999", 0.99999998842592586},
	}

	for _, tt := range excelDates {
		exp, err := time.Parse("2006-01-02T15:04:05.000", tt.expected)
		if err != nil {
			t.Fatalf("Failed to parse expected time %s: %v", tt.expected, err)
		}
		got, err := XldateAsDatetime(tt.xldate, not1904)
		if err != nil {
			t.Errorf("XldateAsDatetime(%f, %v) error = %v", tt.xldate, not1904, err)
			continue
		}
		// Compare with millisecond precision
		diff := got.Sub(exp)
		if diff < 0 {
			diff = -diff
		}
		if diff > time.Millisecond {
			t.Errorf("XldateAsDatetime(%f, %v) = %v, want %v (diff: %v)", tt.xldate, not1904, got, exp, diff)
		}
	}
}
