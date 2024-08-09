package main

import (
	"fmt"
	"log"

	"github.com/tealeg/xlsx"
)

func main() {
	log.Println("Opening originSource.xlsx...")
	originSource, err := openExcelFile("originSource.xlsx")
	if err != nil {
		log.Fatalf("Failed to open originSource.xlsx: %v", err)
	}

	log.Println("Opening newSource.xlsx...")
	newSource, err := openExcelFile("newSource.xlsx")
	if err != nil {
		log.Fatalf("Failed to open newSource.xlsx: %v", err)
	}

	log.Println("Creating new Excel file...")
	newFile, diffSheet, newDataSheet, dataErrorSheet, err := createNewExcelFile()
	if err != nil {
		log.Fatalf("Failed to create new Excel file: %v", err)
	}

	log.Println("Comparing sheets...")
	compareSheets(originSource, newSource, diffSheet, newDataSheet, dataErrorSheet)

	log.Println("Saving output.xlsx...")
	err = saveExcelFile(newFile, "output.xlsx")
	if err != nil {
		log.Fatalf("Failed to save output.xlsx: %v", err)
	}

	fmt.Println("Data saved to output.xlsx")
}

func openExcelFile(filename string) (*xlsx.File, error) {
	return xlsx.OpenFile(filename)
}

func createNewExcelFile() (*xlsx.File, *xlsx.Sheet, *xlsx.Sheet, *xlsx.Sheet, error) {
	headers := []string{
		"id", "province_id", "province", "city_id", "city_type", "city",
		"district_id", "district", "sub_district_id", "sub_district", "postal_code",
		"kode_wilayah_concat", "notes", "kelurahan_found", "kecamatan_found",
		"kode_provinsi", "kode_kota_kabupaten", "kode_kecamatan", "kode_kelurahan", "kode_wilayah",
	}

	newFile := xlsx.NewFile()
	diffSheet, err := newFile.AddSheet("Difference")
	if err != nil {
		return nil, nil, nil, nil, fmt.Errorf("failed to create Difference sheet: %v", err)
	}
	newDataSheet, err := newFile.AddSheet("New Data")
	if err != nil {
		return nil, nil, nil, nil, fmt.Errorf("failed to create New Data sheet: %v", err)
	}
	dataErrorSheet, err := newFile.AddSheet("Data Error")
	if err != nil {
		return nil, nil, nil, nil, fmt.Errorf("failed to create Data Error sheet: %v", err)
	}

	addHeaders(diffSheet, headers)
	addHeaders(newDataSheet, headers)
	addHeaders(dataErrorSheet, headers)

	return newFile, diffSheet, newDataSheet, dataErrorSheet, nil
}

func addHeaders(sheet *xlsx.Sheet, headers []string) {
	headerRow := sheet.AddRow()
	for _, header := range headers {
		headerRow.AddCell().SetString(header)
	}
}

func compareSheets(originSource, newSource *xlsx.File, diffSheet, newDataSheet, errSheet *xlsx.Sheet) {
	for i, sheet2 := range newSource.Sheets {
		if i >= len(originSource.Sheets) {
			break
		}
		sheet1 := originSource.Sheets[i]
		log.Printf("Comparing sheet %d...", i+1)
		compareSheet(sheet1, sheet2, diffSheet, newDataSheet, errSheet, i+1)
	}
}

func compareSheet(sheet1, sheet2 *xlsx.Sheet, diffSheet, newDataSheet, errSheet *xlsx.Sheet, sheetIndex int) {
	maxRows := len(sheet2.Rows)
	for j := 0; j < maxRows; j++ {
		var row1, row2 *xlsx.Row
		row2 = sheet2.Rows[j]
		log.Printf("Comparing row %d in sheet %d...", j+1, sheetIndex)
		for i, row := range sheet1.Rows {
			if row.Cells[0].String() == row2.Cells[0].String() {
				row1 = row
				sheet1.Rows = append(sheet1.Rows[:i], sheet1.Rows[i+1:]...)
				break
			}
		}

		if !validateRow(row2) {
			addRowToSheet(errSheet, row2)
			continue
		}
		if row1 == nil {
			addRowToSheet(newDataSheet, row2)
			continue
		}
		if !compareRows(row1, row2) {
			addRowToSheet(diffSheet, row2)
			continue
		}
	}
}

func compareRows(row1, row2 *xlsx.Row) bool {
	var maxCells int
	if row1 != nil {
		maxCells = len(row1.Cells)
	}
	if row2 != nil && len(row2.Cells) > maxCells {
		maxCells = len(row2.Cells)
	}

	for k := 0; k < maxCells; k++ {
		var cell1, cell2 *xlsx.Cell
		if row1 != nil && k < len(row1.Cells) {
			cell1 = row1.Cells[k]
		}
		if row2 != nil && k < len(row2.Cells) {
			cell2 = row2.Cells[k]
		}
		if cell1.String() != cell2.String() {
			return false
		}
	}
	return true
}

func addRowToSheet(sheet *xlsx.Sheet, row *xlsx.Row) {
	newRow := sheet.AddRow()
	for _, cell := range row.Cells {
		newRow.AddCell().SetString(cell.String())
	}
}

func saveExcelFile(file *xlsx.File, filename string) error {
	return file.Save(filename)
}

func validateRow(row *xlsx.Row) bool {
	// only cell M,N,O optional others are required
	for i, cell := range row.Cells {
		if i == 12 || i == 13 || i == 14 {
			continue
		}
		if cell.String() == "" {
			return false
		}
	}
	return true
}
