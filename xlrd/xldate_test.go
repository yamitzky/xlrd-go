package xlrd

import (
	"math"
	"testing"
)

const datemode = 0 // 1900-based

func TestDateAsTuple(t *testing.T) {
	tests := []struct {
		xldate  float64
		want    [6]int
		wantErr bool
	}{
		{2741., [6]int{1907, 7, 3, 0, 0, 0}, false},
		{38406., [6]int{2005, 2, 23, 0, 0, 0}, false},
		{32266., [6]int{1988, 5, 3, 0, 0, 0}, false},
	}

	for _, tt := range tests {
		year, month, day, hour, minute, second, err := XldateAsTuple(tt.xldate, datemode)
		if (err != nil) != tt.wantErr {
			t.Errorf("XldateAsTuple(%f, %d) error = %v, wantErr %v", tt.xldate, datemode, err, tt.wantErr)
			continue
		}
		if err == nil {
			got := [6]int{year, month, day, hour, minute, second}
			if got != tt.want {
				t.Errorf("XldateAsTuple(%f, %d) = %v, want %v", tt.xldate, datemode, got, tt.want)
			}
		}
	}
}

func TestTimeAsTuple(t *testing.T) {
	tests := []struct {
		xldate  float64
		want    [6]int
		wantErr bool
	}{
		{0.273611, [6]int{0, 0, 0, 6, 34, 0}, false},
		{0.538889, [6]int{0, 0, 0, 12, 56, 0}, false},
		{0.741123, [6]int{0, 0, 0, 17, 47, 13}, false},
	}

	for _, tt := range tests {
		year, month, day, hour, minute, second, err := XldateAsTuple(tt.xldate, datemode)
		if (err != nil) != tt.wantErr {
			t.Errorf("XldateAsTuple(%f, %d) error = %v, wantErr %v", tt.xldate, datemode, err, tt.wantErr)
			continue
		}
		if err == nil {
			got := [6]int{year, month, day, hour, minute, second}
			if got != tt.want {
				t.Errorf("XldateAsTuple(%f, %d) = %v, want %v", tt.xldate, datemode, got, tt.want)
			}
		}
	}
}

func TestXldateFromDateTuple(t *testing.T) {
	tests := []struct {
		year    int
		month   int
		day     int
		want    float64
		wantErr bool
	}{
		{1907, 7, 3, 2741., false},
		{2005, 2, 23, 38406., false},
		{1988, 5, 3, 32266., false},
	}

	for _, tt := range tests {
		got, err := XldateFromDateTuple(tt.year, tt.month, tt.day, datemode)
		if (err != nil) != tt.wantErr {
			t.Errorf("XldateFromDateTuple(%d, %d, %d, %d) error = %v, wantErr %v", tt.year, tt.month, tt.day, datemode, err, tt.wantErr)
			continue
		}
		if err == nil {
			if math.Abs(got-tt.want) > 0.0001 {
				t.Errorf("XldateFromDateTuple(%d, %d, %d, %d) = %f, want %f", tt.year, tt.month, tt.day, datemode, got, tt.want)
			}
		}
	}
}

func TestXldateFromTimeTuple(t *testing.T) {
	tests := []struct {
		hour    int
		minute  int
		second  int
		want    float64
		wantErr bool
	}{
		{6, 34, 0, 0.273611, false},
		{12, 56, 0, 0.538889, false},
		{17, 47, 13, 0.741123, false},
	}

	for _, tt := range tests {
		got, err := XldateFromTimeTuple(tt.hour, tt.minute, tt.second)
		if (err != nil) != tt.wantErr {
			t.Errorf("XldateFromTimeTuple(%d, %d, %d) error = %v, wantErr %v", tt.hour, tt.minute, tt.second, err, tt.wantErr)
			continue
		}
		if err == nil {
			if math.Abs(got-tt.want) > 0.000001 {
				t.Errorf("XldateFromTimeTuple(%d, %d, %d) = %f, want %f", tt.hour, tt.minute, tt.second, got, tt.want)
			}
		}
	}
}

func TestXldateFromDatetimeTuple(t *testing.T) {
	tests := []struct {
		year    int
		month   int
		day     int
		hour    int
		minute  int
		second  int
		want    float64
		wantErr bool
	}{
		{1907, 7, 3, 6, 34, 0, 2741.273611, false},
		{2005, 2, 23, 12, 56, 0, 38406.538889, false},
		{1988, 5, 3, 17, 47, 13, 32266.741123, false},
	}

	for _, tt := range tests {
		got, err := XldateFromDatetimeTuple(tt.year, tt.month, tt.day, tt.hour, tt.minute, tt.second, datemode)
		if (err != nil) != tt.wantErr {
			t.Errorf("XldateFromDatetimeTuple(%d, %d, %d, %d, %d, %d, %d) error = %v, wantErr %v",
				tt.year, tt.month, tt.day, tt.hour, tt.minute, tt.second, datemode, err, tt.wantErr)
			continue
		}
		if err == nil {
			if math.Abs(got-tt.want) > 0.000001 {
				t.Errorf("XldateFromDatetimeTuple(%d, %d, %d, %d, %d, %d, %d) = %f, want %f",
					tt.year, tt.month, tt.day, tt.hour, tt.minute, tt.second, datemode, got, tt.want)
			}
		}
	}
}
