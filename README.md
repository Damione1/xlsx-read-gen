# xlsx-read-gen
## Generate
Use `go run generate/main.go` to generate a file. 

You can define for each column the first row and the fake content generated from the library [gofakeit](https://github.com/brianvoe/gofakeit) with [excelize](https://github.com/qax-os/excelize)

## Read
`go run main.go` to read a xlsx file

You can read a range by defining the range in the `getXLSXParams` function in the format `SheetName!A2:E500`
