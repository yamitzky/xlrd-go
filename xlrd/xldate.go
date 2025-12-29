package xlrd

import (
	"time"
)

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
	// Empty implementation for now
	return 0, 0, 0, 0, 0, 0, nil
}

// XldateAsDatetime converts an Excel number (presumed to represent a date, a datetime or a time)
// into a time.Time value.
//
// xldate: The Excel number
// datemode: 0: 1900-based, 1: 1904-based.
//
// Returns: time.Time value.
func XldateAsDatetime(xldate float64, datemode int) (time.Time, error) {
	// Empty implementation for now
	return time.Time{}, nil
}

// XldateFromDateTuple converts a date tuple to an Excel date number.
func XldateFromDateTuple(year, month, day int, datemode int) (float64, error) {
	// Empty implementation for now
	return 0.0, nil
}

// XldateFromTimeTuple converts a time tuple to an Excel date number.
func XldateFromTimeTuple(hour, minute, second int) (float64, error) {
	// Empty implementation for now
	return 0.0, nil
}

// XldateFromDatetimeTuple converts a datetime tuple to an Excel date number.
func XldateFromDatetimeTuple(year, month, day, hour, minute, second int, datemode int) (float64, error) {
	// Empty implementation for now
	return 0.0, nil
}
