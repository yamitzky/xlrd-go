package xlrd

import (
	"encoding/binary"
	"fmt"
	"unicode/utf16"

	"golang.org/x/text/encoding/charmap"
)

// UnpackString unpacks a string from BIFF data.
func UnpackString(data []byte, pos int, encoding string, lenlen int) (string, error) {
	if pos+lenlen > len(data) {
		return "", fmt.Errorf("insufficient data for string length")
	}

	var nchars int
	if lenlen == 1 {
		nchars = int(data[pos])
	} else {
		nchars = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
	}
	pos += lenlen

	if pos+nchars > len(data) {
		return "", fmt.Errorf("insufficient data for string")
	}

	strBytes := data[pos : pos+nchars]

	// Convert based on encoding
	if encoding == "utf_16_le" {
		if len(strBytes)%2 != 0 {
			return "", fmt.Errorf("invalid UTF-16 string length")
		}
		words := make([]uint16, len(strBytes)/2)
		for i := 0; i < len(words); i++ {
			words[i] = binary.LittleEndian.Uint16(strBytes[i*2 : (i+1)*2])
		}
		return string(utf16.Decode(words)), nil
	}

	// For other encodings, treat as Latin-1 for now
	return string(strBytes), nil
}

// UnpackStringUpdatePos unpacks a string and returns the updated position.
func UnpackStringUpdatePos(data []byte, pos int, encoding string, lenlen int, knownLen *int) (string, int, error) {
	var nchars int
	if knownLen != nil {
		nchars = *knownLen
	} else {
		if pos+lenlen > len(data) {
			return "", pos, fmt.Errorf("insufficient data for string length")
		}
		if lenlen == 1 {
			nchars = int(data[pos])
		} else {
			nchars = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		}
		pos += lenlen
	}

	if pos+nchars > len(data) {
		return "", pos, fmt.Errorf("insufficient data for string")
	}

	strBytes := data[pos : pos+nchars]
	newPos := pos + nchars

	// Convert based on encoding
	if encoding == "utf_16_le" {
		if len(strBytes)%2 != 0 {
			return "", newPos, fmt.Errorf("invalid UTF-16 string length")
		}
		words := make([]uint16, len(strBytes)/2)
		for i := 0; i < len(words); i++ {
			words[i] = binary.LittleEndian.Uint16(strBytes[i*2 : (i+1)*2])
		}
		return string(utf16.Decode(words)), newPos, nil
	}

	// For other encodings, treat as Latin-1 for now
	return string(strBytes), newPos, nil
}

// UnpackUnicode unpacks a Unicode string from BIFF data.
func UnpackUnicode(data []byte, pos int, lenlen int) (string, error) {
	if pos+lenlen > len(data) {
		return "", fmt.Errorf("insufficient data for unicode length")
	}

	var nchars int
	if lenlen == 1 {
		nchars = int(data[pos])
	} else {
		nchars = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
	}
	pos += lenlen

	if nchars == 0 {
		return "", nil
	}

	if pos >= len(data) {
		return "", fmt.Errorf("insufficient data for unicode options")
	}

	options := data[pos]
	pos++

	// Handle richtext and phonetic flags
	if options&0x08 != 0 {
		// richtext
		if pos+2 > len(data) {
			return "", fmt.Errorf("insufficient data for richtext")
		}
		pos += 2
	}
	if options&0x04 != 0 {
		// phonetic
		if pos+4 > len(data) {
			return "", fmt.Errorf("insufficient data for phonetic")
		}
		pos += 4
	}

	if options&0x01 != 0 {
		// Uncompressed UTF-16-LE
		if pos+2*nchars > len(data) {
			return "", fmt.Errorf("insufficient data for UTF-16 string")
		}
		rawstrg := data[pos : pos+2*nchars]
		words := make([]uint16, nchars)
		for i := 0; i < nchars; i++ {
			words[i] = binary.LittleEndian.Uint16(rawstrg[i*2 : (i+1)*2])
		}
		return string(utf16.Decode(words)), nil
	} else {
		// Compressed (Latin-1)
		if pos+nchars > len(data) {
			return "", fmt.Errorf("insufficient data for compressed string")
		}
		latin1Bytes := data[pos : pos+nchars]
		utf8Bytes, err := charmap.ISO8859_1.NewDecoder().Bytes(latin1Bytes)
		if err != nil {
			return "", fmt.Errorf("failed to decode Latin-1: %v", err)
		}
		return string(utf8Bytes), nil
	}
}

// UnpackUnicodeUpdatePos unpacks a Unicode string and returns the updated position.
func UnpackUnicodeUpdatePos(data []byte, pos int, lenlen int, knownLen *int) (string, int, error) {
	var nchars int
	if knownLen != nil {
		nchars = *knownLen
	} else {
		if pos+lenlen > len(data) {
			return "", pos, fmt.Errorf("insufficient data for unicode length")
		}
		if lenlen == 1 {
			nchars = int(data[pos])
		} else {
			nchars = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		}
		pos += lenlen
	}

	if nchars == 0 {
		if pos >= len(data) {
			return "", pos, nil
		}
		return "", pos, nil
	}

	if pos >= len(data) {
		return "", pos, fmt.Errorf("insufficient data for unicode options")
	}

	options := data[pos]
	pos++

	phonetic := (options & 0x04) != 0
	richtext := (options & 0x08) != 0

	if richtext {
		if pos+2 > len(data) {
			return "", pos, fmt.Errorf("insufficient data for richtext")
		}
		pos += 2
	}
	if phonetic {
		if pos+4 > len(data) {
			return "", pos, fmt.Errorf("insufficient data for phonetic")
		}
		pos += 4
	}

	var str string
	if options&0x01 != 0 {
		// Uncompressed UTF-16-LE
		if pos+2*nchars > len(data) {
			return "", pos, fmt.Errorf("insufficient data for UTF-16 string")
		}
		rawstrg := data[pos : pos+2*nchars]
		words := make([]uint16, nchars)
		for i := 0; i < nchars; i++ {
			words[i] = binary.LittleEndian.Uint16(rawstrg[i*2 : (i+1)*2])
		}
		str = string(utf16.Decode(words))
		pos += 2 * nchars
	} else {
		// Compressed (Latin-1)
		if pos+nchars > len(data) {
			return "", pos, fmt.Errorf("insufficient data for compressed string")
		}
		str = string(data[pos : pos+nchars])
		pos += nchars
	}

	if richtext {
		// Skip richtext data (would need to read rt count)
		// For now, skip
	}
	if phonetic {
		// Skip phonetic data (would need to read sz)
		// For now, skip
	}

	return str, pos, nil
}
