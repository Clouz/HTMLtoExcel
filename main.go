package main

import (
	"fmt"
	"log"
	"os"
	"path"
	"strings"

	"github.com/xuri/excelize"
	"golang.org/x/net/html"
	"golang.org/x/text/encoding/charmap"
)

func main() {
	log.Println("Apro il file...")

	f, err := os.Open(os.Args[1])
	if err != nil {
		log.Fatal(err)
	}

	file := charmap.Windows1252.NewDecoder().Reader(f)
	log.Println("Codifica caratteri da Windows-1252 a UTF-8")

	z := html.NewTokenizer(file)

	data := make(map[int][][]string)
	dataIndex := 0
	dataRow := []string{}
	dataRowCount := 0

	table := 0
	row := 0
	cell := 0

mainLoop:
	for {
		tt := z.Next()
		switch tt {
		case html.ErrorToken:
			log.Println(z.Err())
			break mainLoop
		case html.TextToken:
			if cell > 0 {
				s := string(z.Text())
				//fmt.Println(s)
				dataRow = append(dataRow, s)
			}

		case html.StartTagToken, html.EndTagToken:
			tn, _ := z.TagName()
			//fmt.Println(tt)

			switch {
			case strings.Contains(string(tn), "table"):
				if tt == html.StartTagToken {
					table++

					dataIndex++
					//log.Println("[TABLE] ++")
				} else {
					table--
					//log.Println("[TABLE] --")
				}

			case strings.Contains(string(tn), "tr") && table > 0:
				if tt == html.StartTagToken {
					//log.Println("[row] ++")
					row++
				} else {
					//log.Println("[row] --")
					row--
					data[dataIndex] = append(data[dataIndex], dataRow)
					dataRow = []string{}
					dataRowCount = 0
					//fmt.Println()
				}

			case (strings.Contains(string(tn), "td") || strings.Contains(string(tn), "th")) && row > 0:
				if tt == html.StartTagToken {
					//log.Println("[cell] ++")
					cell++
					dataRowCount++
				} else {
					//log.Println("[cell] -- ", dataRowCount, " ", len(dataRow))
					if dataRowCount > len(dataRow) {
						dataRow = append(dataRow, "")
					}
					cell--
				}
			}
		}
	}

	log.Println("Lettura file completata...")
	log.Println("Preparazione nuovo Excel...")

	xlsx := excelize.NewFile()
	for index := 1; index <= len(data); index++ {
		table := fmt.Sprint("Sheet", index)
		if index > 1 {
			xlsx.NewSheet(table)
		}
		for row := 0; row < len(data[index]); row++ {
			for col := 0; col < len(data[index][row]); col++ {
				ax := fmt.Sprint(excelize.ToAlphaString(col), row+1)
				xlsx.SetCellValue(table, ax, data[index][row][col])
				//fmt.Println(ax + " " + data[index][row][col])
			}
			//fmt.Println()
		}

	}

	filename := strings.Trim(path.Base(os.Args[1]), path.Ext(os.Args[1]))
	log.Println("Salvataggio file", filename)

	err = xlsx.SaveAs(fmt.Sprint("./", filename, ".xlsx"))
	if err != nil {

	}
	log.Println("Operazione completata, premere un tasto per chiudere...")
	fmt.Scanln()
}
