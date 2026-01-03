package xls

import (
	"bytes"
	"encoding/binary"
	"math"
	"testing"
)

func TestFormulaStringCachedResultBIFF8(t *testing.T) {
	biff := buildFormulaStringBiff8()
	wb := newWorkBookFromOle2(bytes.NewReader(biff))
	sheet := wb.GetSheet(0)
	if sheet == nil {
		t.Fatal("expected sheet 0")
	}
	row := sheet.Row(0)
	if row == nil {
		t.Fatal("expected row 0")
	}
	if got := row.Col(0); got != "Normal" {
		t.Fatalf("labelSST cell mismatch: got %q", got)
	}
	if got := row.Col(1); got != "FormulaString" {
		t.Fatalf("formula string cell mismatch: got %q", got)
	}
	if got := row.Col(2); got != "" {
		t.Fatalf("numeric formula cell changed: got %q", got)
	}
}

func buildFormulaStringBiff8() []byte {
	bofGlobals := record(0x0809, biff8BOFData(0x0005))
	sst := record(0x00FC, sstData([]string{"Normal"}))
	eof := record(0x000A, nil)

	bs := record(0x0085, boundsheetData(0, "Sheet1"))
	sheetOffset := len(bofGlobals) + len(bs) + len(sst) + len(eof)
	bs = record(0x0085, boundsheetData(uint32(sheetOffset), "Sheet1"))

	var globals []byte
	globals = append(globals, bofGlobals...)
	globals = append(globals, bs...)
	globals = append(globals, sst...)
	globals = append(globals, eof...)

	sheet := buildFormulaStringSheet()
	return append(globals, sheet...)
}

func buildFormulaStringSheet() []byte {
	var sheet []byte
	sheet = append(sheet, record(0x0809, biff8BOFData(0x0010))...)
	sheet = append(sheet, record(0x00FD, labelSSTData(0, 0, 0, 0))...)

	stringResult := [8]byte{0xff, 0xff, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}
	sheet = append(sheet, record(0x0006, formulaData(0, 1, 0, stringResult))...)
	sheet = append(sheet, record(0x0207, stringRecordData("FormulaString"))...)

	var numberResult [8]byte
	binary.LittleEndian.PutUint64(numberResult[:], math.Float64bits(123.45))
	sheet = append(sheet, record(0x0006, formulaData(0, 2, 0, numberResult))...)

	sheet = append(sheet, record(0x000A, nil)...)
	return sheet
}

func record(id uint16, data []byte) []byte {
	out := make([]byte, 4+len(data))
	binary.LittleEndian.PutUint16(out[0:2], id)
	binary.LittleEndian.PutUint16(out[2:4], uint16(len(data)))
	copy(out[4:], data)
	return out
}

func biff8BOFData(bofType uint16) []byte {
	var buf bytes.Buffer
	header := biffHeader{
		Ver:  0x0600,
		Type: bofType,
	}
	binary.Write(&buf, binary.LittleEndian, &header)
	return buf.Bytes()
}

func boundsheetData(sheetOffset uint32, name string) []byte {
	var buf bytes.Buffer
	binary.Write(&buf, binary.LittleEndian, sheetOffset)
	buf.WriteByte(0x00) // visibility
	buf.WriteByte(0x00) // worksheet
	buf.WriteByte(byte(len(name)))
	buf.Write(biff8StringData(name))
	return buf.Bytes()
}

func sstData(strings []string) []byte {
	var buf bytes.Buffer
	info := SstInfo{Total: uint32(len(strings)), Count: uint32(len(strings))}
	binary.Write(&buf, binary.LittleEndian, &info)
	for _, s := range strings {
		binary.Write(&buf, binary.LittleEndian, uint16(len(s)))
		buf.Write(biff8StringData(s))
	}
	return buf.Bytes()
}

func biff8StringData(s string) []byte {
	var buf bytes.Buffer
	buf.WriteByte(0x00) // compressed 8-bit
	buf.WriteString(s)
	return buf.Bytes()
}

func labelSSTData(row, col, xf uint16, sst uint32) []byte {
	var buf bytes.Buffer
	binary.Write(&buf, binary.LittleEndian, row)
	binary.Write(&buf, binary.LittleEndian, col)
	binary.Write(&buf, binary.LittleEndian, xf)
	binary.Write(&buf, binary.LittleEndian, sst)
	return buf.Bytes()
}

func formulaData(row, col, xf uint16, result [8]byte) []byte {
	var buf bytes.Buffer
	binary.Write(&buf, binary.LittleEndian, row)
	binary.Write(&buf, binary.LittleEndian, col)
	binary.Write(&buf, binary.LittleEndian, xf)
	buf.Write(result[:])
	binary.Write(&buf, binary.LittleEndian, uint16(0)) // flags
	binary.Write(&buf, binary.LittleEndian, uint32(0)) // reserved / unused
	return buf.Bytes()
}

func stringRecordData(s string) []byte {
	var buf bytes.Buffer
	binary.Write(&buf, binary.LittleEndian, uint16(len(s)))
	buf.Write(biff8StringData(s))
	return buf.Bytes()
}
