package main

import (
	"fmt"
	"regexp"
	"strings"

	excelize "github.com/xuri/excelize/v2"
)

type ingestionParameters struct {
	sheetName    string
	cellFromName string
	cellToName   string
	cellFromAxis axis
	cellToAxis   axis
	isRange      bool
}

type axis struct {
	x int
	y int
}

func main() {
	xlsxFile, err := excelize.OpenFile("test_import_1000_rows.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := xlsxFile.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	ingestionParams, err := getXLSXParams("", xlsxFile)
	if err != nil {
		fmt.Println(err)
		return
	}

	rows, err := xlsxFile.Rows(ingestionParams.sheetName)
	if err != nil {
		fmt.Println("error while getting Rows:", err)
		return
	}
	y := 1
	for ; rows.Next(); y++ {
		if !ingestionParams.isRange || y >= ingestionParams.cellFromAxis.y && y <= ingestionParams.cellToAxis.y {
			row, err := rows.Columns()
			if err != nil {
				fmt.Println("error in Columns:s", err)
				return
			}
			ingestedRow := row
			if ingestionParams.isRange {
				ingestedRow = row[ingestionParams.cellFromAxis.x-1 : ingestionParams.cellToAxis.x]
			}

			if len(ingestedRow) == 0 {
				fmt.Println("empty row")
				return
			}

			fmt.Println(ingestedRow)

		}
	}
	if !rows.Next() && y == 1 {
		fmt.Println("empty file")
		return
	}

}

func getXLSXParams(paramString string, xlsxFile *excelize.File) (*ingestionParameters, error) {
	ingestionParams := ingestionParameters{
		isRange: false,
	}

	if len(paramString) == 0 {
		ingestionParams.sheetName = xlsxFile.GetSheetName(0)
		return &ingestionParams, nil
	}
	ingestionParams.sheetName = paramString

	if strings.Contains(paramString, "!") && strings.Contains(paramString, ":") {
		ingestionParams.isRange = false
		regex := regexp.MustCompile(`^([a-zA-Z0-9]+)!([a-zA-Z]+[0-9]+):([a-zA-Z]+[0-9]+)$`)
		matches := regex.FindStringSubmatch(paramString)
		if len(matches) != 4 {
			return nil, fmt.Errorf("could not parse xlsx range %s", paramString)
		}
		ingestionParams.sheetName = matches[1]
		x, y, err := excelize.CellNameToCoordinates(matches[2])
		if err != nil {
			return nil, err
		}
		ingestionParams.cellFromAxis = axis{
			x: x,
			y: y,
		}
		x, y, err = excelize.CellNameToCoordinates(matches[3])
		if err != nil {
			return nil, err
		}
		ingestionParams.cellToAxis = axis{
			x: x,
			y: y,
		}
	}
	return &ingestionParams, nil
}

func validateXLSXParams(ingestionParams *ingestionParameters) error {
	if ingestionParams.isRange {
		if ingestionParams.cellFromAxis.x > ingestionParams.cellToAxis.x {
			return fmt.Errorf("invalid xlsx range, from column is greater than to column")
		}
		if ingestionParams.cellFromAxis.y > ingestionParams.cellToAxis.y {
			return fmt.Errorf("invalid xlsx range, from row is greater than to row")
		}
	}

	return nil
}
