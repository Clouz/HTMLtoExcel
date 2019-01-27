package main

import (
	"fmt"
	"log"
	"os"
	"strings"

	"golang.org/x/net/html"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	log.Println("Start...")

	xlsx := excelize.NewFile()
	xlsx.SetCellValue("Sheet1", "A2", "Hello world.")
	err := xlsx.SaveAs("./Analisi.xlsx")
	if err != nil {
		fmt.Println(err)
	}

	file, err := os.Open(os.Args[1])
	if err != nil {
		log.Fatal(err)
	}

	z := html.NewTokenizer(file)

	table := 0
	row := 0
	cell := 0

	for {
		tt := z.Next()
		switch tt {
		case html.ErrorToken:
			log.Println(z.Err())
			return
		case html.TextToken:
			if cell > 0 {
				fmt.Print(string(z.Text()), "\t")
			}

		case html.StartTagToken, html.EndTagToken:
			tn, _ := z.TagName()

			switch {
			case strings.Contains(string(tn), "table"):
				if tt == html.StartTagToken {
					table++
					log.Println("[TABLE] ++")
				} else {
					table--
					log.Println("[TABLE] --")
				}

			case strings.Contains(string(tn), "tr") && table > 0:
				if tt == html.StartTagToken {
					row++
				} else {
					row--
					fmt.Println("riga")

				}

			case strings.Contains(string(tn), "td") && row > 0:
				if tt == html.StartTagToken {
					cell++

				} else {
					cell--

				}

			}

		}
	}

}
