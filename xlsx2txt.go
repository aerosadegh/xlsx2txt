package main

import (
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func printString(n int) string {
	var arr [10000]int
	i := 0
	st := ""
	// Step 1: Converting to number
	// assuming 0 in number system

	for n > 0 {
		arr[i] = n % 26
		n = int(n / 26)
		i++
	}

	// Step 2: Getting rid of 0, as 0 is
	// not part of number system
	for j := 0; j < i-1; j++ {
		if arr[j] <= 0 {
			arr[j] += 26
			arr[j+1] = arr[j+1] - 1
		}
	}

	for k := i; k > -1; k-- {
		if arr[k] > 0 {
			st = st + string(int('A')+(arr[k]-1))
		}
	}
	return st
}

func main() {
	argsWithProg := os.Args
	filename := "Book1.xlsx"

	if len(argsWithProg) > 1 {
		filename = argsWithProg[1]
	}

	f, err := excelize.OpenFile(filename)
	if err != nil {
		fmt.Println(err)
		return
	}
	shl := f.GetSheetList()

	// Get all the rows in the Sheet1.
	for _, sheet := range shl {
		fmt.Println("SHEET:<<< ", sheet, " >>>")
		rows, _ := f.GetRows(sheet)
		for i, row := range rows {
			fmt.Print(sheet, ":> ")
			for j, colCell := range row {
				fmt.Print("|:", printString(j+1), i+1, ":> ", colCell, "  \t")
			}
			fmt.Println()
		}
	}

}
