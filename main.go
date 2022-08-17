package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()
	sheetName := "Transcript"
	f.SetSheetName("Sheet1", sheetName)
	data := [][]interface{}{
		{"Student Exam Score"},
		{"Type : Mid term Exam", nil, nil, nil, "Core Curriculum", nil, nil, "Science"},
		{"Number", "ID", "Name", "Class", "Language Arts", "Mathemathics", "History", "Chemistry", "Biology", "Physics", "Total"},
		{1, 1001, "Student A", "Class 1", 80, 90, 80, 90, 70, 80},
		{2, 1002, "Student B", "Class 2", 100, 90, 80, 90, 100, 80},
		{3, 1003, "Student C", "Class 3", 100, 90, 80, 90, 100, 80},
		{4, 1004, "Student D", "Class 1", 70, 90, 80, 90, 100, 80},
		{5, 1005, "Student E", "Class 3", 100, 90, 80, 90, 100, 80},
		{6, 1006, "Student F", "Class 2", 100, 90, 80, 90, 60, 80},
		{7, 1007, "Student G", "Class 1", 100, 90, 80, 90, 100, 80},
	}

	// looping datanya cuy
	for i, row := range data {
		startCell, err := excelize.JoinCellName("A", i+1)
		if err != nil {
			fmt.Println(err)
			return
		}
		if err := f.SetSheetRow(sheetName, startCell, &row); err != nil {
			fmt.Println(err)
			return
		}
	}
	// ini buat ngsum pke rumus
	formulaType, reff := excelize.STCellFormulaTypeShared, "K4:K10"
	if err := f.SetCellFormula(sheetName, "K4", "=SUM(E4:J4)", excelize.FormulaOpts{Ref: &reff, Type: &formulaType}); err != nil {
		fmt.Println(err)
		return
	}
	// end

	// ng merge cell di excel
	mergeCellsRange := [][]string{{"A1", "K1"}, {"A2", "D2"}, {"E2", "G2"}, {"H2", "J2"}}
	for _, ranges := range mergeCellsRange {
		if err := f.MergeCell(sheetName, ranges[0], ranges[1]); err != nil {
			fmt.Println(err)
			return
		}
	}
	// end

	// make style center
	style1, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#DFEBF6"}, Pattern: 1},
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	if f.SetCellStyle(sheetName, "A1", "A1", style1); err != nil {
		fmt.Println(err)
		return
	}
	style2, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	for _, cell := range []string{"A2", "H2", "E2"} {
		if err := f.SetCellStyle(sheetName, cell, cell, style2); err != nil {
			fmt.Println(err)
			return
		}
	}
	// end

	// mengatur width column dari cell D - K
	if err := f.SetColWidth(sheetName, "D", "K", 14); err != nil {
		fmt.Println(err)
		return
	}
	// end

	if err := f.AddTable(sheetName, "A3", "K10", `{
        "table_name" : "table",
        "table_style" : "TableStyleLight2"
    }`); err != nil {
		fmt.Println(err)
		return
	}

	// ngsave excelnya
	if err := f.SaveAs("New-Excel.xlsx"); err != nil {
		fmt.Println(err)
	}
}
