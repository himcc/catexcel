package main

import (
	"fmt"
	"os"
	"strconv"

	"github.com/mkideal/cli"
	"github.com/xuri/excelize/v2"
)

type Params struct {
	cli.Helper

	All    bool     `cli:"a,all" usage:"output all sheets" dft:"false"`
	Format string   `cli:"f,format" usage:"tsv or csv" dft:"tsv"`
	Sheet  []string `cli:"s,sheet" usage:"sheet name or index (from 0) for output"`
}

type MyError struct {
	msg string
}

func (t MyError) Error() string {
	return t.msg
}

var (
	CmdDescs []string = []string{
		"cat excel file",
		"catexcel a.xlsx b.xlsx               # show top 10 lines for every sheet",
		"catexcel -s 0 -s 1 a.xlsx            # show complete content in first sheet and second sheet",
		"catexcel -f csv -s students a.xlsx   # show complete content in students sheet with csv format(default tsv)",
	}
)

func main() {

	os.Exit(cli.Run(new(Params), func(ctx *cli.Context) error {
		params := ctx.Argv().(*Params)
		fileList := ctx.Args()

		// check file
		if len(fileList) < 1 {
			fmt.Fprintln(os.Stderr, "One or more files need to be specified.")
			fmt.Fprintln(os.Stderr, "")
			fmt.Fprintln(os.Stderr, ctx.Usage())
			return MyError{msg: ""}
		}

		// check format
		if params.Format != "csv" && params.Format != "tsv" {
			fmt.Fprintln(os.Stderr, "format must be csv or tsv")
			fmt.Fprintln(os.Stderr, "")
			fmt.Fprintln(os.Stderr, ctx.Usage())
			return MyError{msg: ""}
		}

		if !params.All && len(params.Sheet) < 1 {
			showTop(fileList, params)
		} else if params.All {
			show(fileList, params, getSheetFilter([]string{}))
		} else if len(params.Sheet) > 0 {
			show(fileList, params, getSheetFilter(params.Sheet))
		}

		return nil
	}, CmdDescs...))

}

func show(files []string, p *Params, fn func(int, string) bool) {
	showf := func(filename string, p *Params, fn func(int, string) bool) {
		f, err := excelize.OpenFile(filename)
		if err != nil {
			fmt.Fprintln(os.Stderr, err)
			return
		}
		defer func() {
			// Close the spreadsheet.
			if err := f.Close(); err != nil {
				fmt.Fprintln(os.Stderr, err)
			}
		}()

		separator := "\t"
		if "csv" == p.Format {
			separator = ","
		}

		sheetList := f.GetSheetList()

		for i, sheet := range sheetList {
			if !fn(i, sheet) {
				continue
			}
			rows, err := f.Rows(sheet)
			if err != nil {
				fmt.Fprintln(os.Stderr, err)
				return
			}
			for rows.Next() {
				row, err := rows.Columns()
				if err != nil {
					fmt.Fprintln(os.Stderr, err)
				}
				for colNo, colCell := range row {
					if colNo != 0 {
						fmt.Print(separator)
					}
					fmt.Print(colCell)
				}
				fmt.Println()
			}
			if err = rows.Close(); err != nil {
				fmt.Fprintln(os.Stderr, err)
			}
		}

	}
	for _, filename := range files {
		showf(filename, p, fn)
	}
}

func getSheetFilter(sheets []string) func(int, string) bool {
	if len(sheets) < 1 {
		return func(i int, s string) bool {
			return true
		}
	} else {
		nameMap := make(map[string]int)
		noMap := make(map[int]int)
		for _, s := range sheets {
			nameMap[s] = 1
			num, e := strconv.Atoi(s)
			if e == nil {
				noMap[num] = 1
			}
		}
		return func(sheetNo int, sheetName string) bool {
			if _, ok := nameMap[sheetName]; ok {
				return true
			}
			if _, ok := noMap[sheetNo]; ok {
				return true
			}
			return false
		}
	}
}

func showTop(files []string, p *Params) {

	showf := func(filename string, showFileName bool, p *Params, fileNo int) {
		f, err := excelize.OpenFile(filename)
		if err != nil {
			fmt.Fprintln(os.Stderr, err)
			return
		}
		defer func() {
			// Close the spreadsheet.
			if err := f.Close(); err != nil {
				fmt.Fprintln(os.Stderr, err)
			}
		}()

		separator := "\t"
		if "csv" == p.Format {
			separator = ","
		}

		sheetList := f.GetSheetList()

		for i, sheet := range sheetList {
			if i == 0 && fileNo == 0 {
			} else {
				fmt.Println("")
			}
			if showFileName {
				// fmt.Println(">>" + filename)
				fmt.Println(">>", filename)
			}
			// fmt.Println(">>" + strconv.Itoa(i) + ":" + sheet)
			fmt.Println(">>", i, ":", sheet)

			rows, err := f.Rows(sheet)
			if err != nil {
				fmt.Fprintln(os.Stderr, err)
				return
			}
			rowNo := 1
			for rows.Next() {
				if rowNo > 10 {
					break
				} else {
					rowNo++
				}
				row, err := rows.Columns()
				if err != nil {
					fmt.Fprintln(os.Stderr, err)
				}
				for colNo, colCell := range row {
					if colNo != 0 {
						fmt.Print(separator)
					}
					fmt.Print(colCell)
				}
				fmt.Println()
			}
			if err = rows.Close(); err != nil {
				fmt.Fprintln(os.Stderr, err)
			}
		}

	}

	showFileName := len(files) > 1
	for i, filePath := range files {
		showf(filePath, showFileName, p, i)

	}
}
