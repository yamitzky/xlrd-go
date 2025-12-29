package xlrd

import (
	"fmt"
	"math"
	"time"
)

var (
	jdnDelta = [2]int{2415080 - 61, 2416482 - 1}
)

const (
	xldaysTooLarge1900 = 2958466
	xldaysTooLarge1904 = 2958466 - 1462
)

var (
	epoch1904        = time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)
	epoch1900        = time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)
	epoch1900Minus1  = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
)

var daysInMonth = [13]int{0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}

// XLDateError is the base type for all datetime-related errors.
type XLDateError struct {
	Message string
}

func (e *XLDateError) Error() string {
	return e.Message
}

// XLDateNegative indicates that xldate < 0.00
type XLDateNegative struct {
	XLDateError
}

// XLDateAmbiguous indicates the 1900 leap-year problem (datemode == 0 and 1.0 <= xldate < 61.0)
type XLDateAmbiguous struct {
	XLDateError
}

// XLDateTooLarge indicates Gregorian year 10000 or later
type XLDateTooLarge struct {
	XLDateError
}

// XLDateBadDatemode indicates that datemode arg is neither 0 nor 1
type XLDateBadDatemode struct {
	XLDateError
}

// XLDateBadTuple indicates a bad tuple parameter
type XLDateBadTuple struct {
	XLDateError
}

// leap returns 1 if year is a leap year, 0 otherwise.
func leap(y int) int {
	if y%4 != 0 {
		return 0
	}
	if y%100 != 0 {
		return 1
	}
	if y%400 != 0 {
		return 0
	}
	return 1
}

// XldateAsTuple converts an Excel number (presumed to represent a date, a datetime or a time)
// into a tuple suitable for feeding to datetime constructors.
//
// xldate: The Excel number
// datemode: 0: 1900-based, 1: 1904-based.
//
// Returns: Gregorian (year, month, day, hour, minute, nearest_second).
//
// Special case: If 0.0 <= xldate < 1.0, it is assumed to represent a time;
// (0, 0, 0, hour, minute, second) will be returned.
func XldateAsTuple(xldate float64, datemode int) (int, int, int, int, int, int, error) {
	if datemode != 0 && datemode != 1 {
		return 0, 0, 0, 0, 0, 0, &XLDateBadDatemode{XLDateError{Message: fmt.Sprintf("Invalid datemode: %d", datemode)}}
	}
	if xldate == 0.00 {
		return 0, 0, 0, 0, 0, 0, nil
	}
	if xldate < 0.00 {
		return 0, 0, 0, 0, 0, 0, &XLDateNegative{XLDateError{Message: fmt.Sprintf("xldate < 0.00: %f", xldate)}}
	}
	xldays := int(xldate)
	frac := xldate - float64(xldays)
	seconds := int(math.Round(frac * 86400.0))
	if seconds < 0 || seconds > 86400 {
		return 0, 0, 0, 0, 0, 0, &XLDateError{Message: fmt.Sprintf("Invalid seconds: %d", seconds)}
	}
	
	hour := 0
	minute := 0
	second := 0
	
	if seconds == 86400 {
		hour = 0
		minute = 0
		second = 0
		xldays++
	} else {
		minutes := seconds / 60
		second = seconds % 60
		hour = minutes / 60
		minute = minutes % 60
	}
	
	xldaysTooLarge := xldaysTooLarge1900
	if datemode == 1 {
		xldaysTooLarge = xldaysTooLarge1904
	}
	if xldays >= xldaysTooLarge {
		return 0, 0, 0, 0, 0, 0, &XLDateTooLarge{XLDateError{Message: fmt.Sprintf("xldate too large: %f", xldate)}}
	}
	
	if xldays == 0 {
		return 0, 0, 0, hour, minute, second, nil
	}
	
	if xldays < 61 && datemode == 0 {
		return 0, 0, 0, 0, 0, 0, &XLDateAmbiguous{XLDateError{Message: fmt.Sprintf("1900 leap-year problem: %f", xldate)}}
	}
	
	jdn := xldays + jdnDelta[datemode]
	yreg := ((((jdn*4+274277)/146097)*3/4)+jdn+1363)*4 + 3
	mp := ((yreg % 1461) / 4) * 535 + 333
	d := ((mp % 16384) / 535) + 1
	mp >>= 14
	if mp >= 10 {
		return ((yreg / 1461) - 4715), (mp - 9), d, hour, minute, second, nil
	}
	return ((yreg / 1461) - 4716), (mp + 3), d, hour, minute, second, nil
}

// XldateAsDatetime converts an Excel number (presumed to represent a date, a datetime or a time)
// into a time.Time value.
//
// xldate: The Excel number
// datemode: 0: 1900-based, 1: 1904-based.
//
// Returns: time.Time value.
func XldateAsDatetime(xldate float64, datemode int) (time.Time, error) {
	var epoch time.Time
	if datemode == 1 {
		epoch = epoch1904
	} else {
		if xldate < 60 {
			epoch = epoch1900
		} else {
			// Workaround Excel 1900 leap year bug by adjusting the epoch.
			epoch = epoch1900Minus1
		}
	}
	
	days := int(xldate)
	fraction := xldate - float64(days)
	
	// Get the integer and decimal seconds in Excel's millisecond resolution.
	seconds := int(math.Round(fraction * 86400000.0))
	secs := seconds / 1000
	milliseconds := seconds % 1000
	
	return epoch.AddDate(0, 0, days).Add(time.Duration(secs)*time.Second + time.Duration(milliseconds)*time.Millisecond), nil
}

// XldateFromDateTuple converts a date tuple to an Excel date number.
func XldateFromDateTuple(year, month, day int, datemode int) (float64, error) {
	if datemode != 0 && datemode != 1 {
		return 0.0, &XLDateBadDatemode{XLDateError{Message: fmt.Sprintf("Invalid datemode: %d", datemode)}}
	}
	
	if year == 0 && month == 0 && day == 0 {
		return 0.00, nil
	}
	
	if year < 1900 || year > 9999 {
		return 0.0, &XLDateBadTuple{XLDateError{Message: fmt.Sprintf("Invalid year: (%d, %d, %d)", year, month, day)}}
	}
	if month < 1 || month > 12 {
		return 0.0, &XLDateBadTuple{XLDateError{Message: fmt.Sprintf("Invalid month: (%d, %d, %d)", year, month, day)}}
	}
	maxDay := daysInMonth[month]
	if month == 2 && leap(year) == 1 {
		maxDay = 29
	}
	if day < 1 || day > maxDay {
		return 0.0, &XLDateBadTuple{XLDateError{Message: fmt.Sprintf("Invalid day: (%d, %d, %d)", year, month, day)}}
	}
	
	Yp := year + 4716
	M := month
	var Mp int
	if M <= 2 {
		Yp = Yp - 1
		Mp = M + 9
	} else {
		Mp = M - 3
	}
	jdn := (1461*Yp/4) + ((979*Mp+16)/32) + day - 1364 - (((Yp+184)/100)*3/4)
	xldays := jdn - jdnDelta[datemode]
	if xldays <= 0 {
		return 0.0, &XLDateBadTuple{XLDateError{Message: fmt.Sprintf("Invalid (year, month, day): (%d, %d, %d)", year, month, day)}}
	}
	if xldays < 61 && datemode == 0 {
		return 0.0, &XLDateAmbiguous{XLDateError{Message: fmt.Sprintf("Before 1900-03-01: (%d, %d, %d)", year, month, day)}}
	}
	return float64(xldays), nil
}

// XldateFromTimeTuple converts a time tuple to an Excel date number.
func XldateFromTimeTuple(hour, minute, second int) (float64, error) {
	if hour < 0 || hour >= 24 || minute < 0 || minute >= 60 || second < 0 || second >= 60 {
		return 0.0, &XLDateBadTuple{XLDateError{Message: fmt.Sprintf("Invalid (hour, minute, second): (%d, %d, %d)", hour, minute, second)}}
	}
	return ((float64(second)/60.0 + float64(minute)) / 60.0 + float64(hour)) / 24.0, nil
}

// XldateFromDatetimeTuple converts a datetime tuple to an Excel date number.
func XldateFromDatetimeTuple(year, month, day, hour, minute, second int, datemode int) (float64, error) {
	datePart, err := XldateFromDateTuple(year, month, day, datemode)
	if err != nil {
		return 0.0, err
	}
	timePart, err := XldateFromTimeTuple(hour, minute, second)
	if err != nil {
		return 0.0, err
	}
	return datePart + timePart, nil
}
