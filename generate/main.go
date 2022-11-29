package main

import (
	"fmt"

	"github.com/brianvoe/gofakeit/v6"
	excelize "github.com/xuri/excelize/v2"
)

type column struct {
	name    string
	content interface{}
}

func main() {
	fileName := "test_import_%d_rows.xlsx"
	firstPageName := "Sheet1"
	numberOfRows := 1000

	f := excelize.NewFile()
	index := f.NewSheet(firstPageName)

	for y := 1; y < numberOfRows+2; y++ {

		columns := []column{
			{
				name:    "email",
				content: gofakeit.Email(),
			},
			{
				name:    "ip",
				content: gofakeit.IPv4Address(),
			},
			{
				name:    "phone",
				content: gofakeit.Phone(),
			},
			{
				name:    "gender",
				content: gofakeit.Gender(),
			},
			{
				name:    "currency",
				content: gofakeit.CurrencyShort(),
			},
			{
				name:    "petname",
				content: gofakeit.PetName(),
			},
		}

		for x, column := range columns {

			cellName, err := excelize.CoordinatesToCellName(x+1, y)
			if err != nil {
				fmt.Println("error in CoordinatesToCellName", err)
				return
			}
			content := gofakeit.Generate(column.content.(string))
			if y == 1 {
				content = column.name
			}
			f.SetCellValue(firstPageName, cellName, content)

		}
	}

	f.SetActiveSheet(index)
	if err := f.SaveAs(fmt.Sprintf(fileName, numberOfRows)); err != nil {
		fmt.Println(err)
	}

}
