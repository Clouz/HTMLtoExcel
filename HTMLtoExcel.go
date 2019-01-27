//go:generate goversioninfo
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

	if len(os.Args) != 2 {
		log.Println("Nessun file selezionato, il programma verrà chiuso")
		fmt.Scanln()
		os.Exit(0)
	}

	ext := strings.ToLower(path.Ext(os.Args[1]))
	if ext != ".html" && ext != ".htm" {
		log.Println("Il file selezionato non è valito (", ext, "). Gli unici file validi devono avere estensione .html o .htm")
		fmt.Scanln()
		os.Exit(0)
	}

	// apro il file passato come primo argomento
	f, err := os.Open(os.Args[1])
	if err != nil {
		log.Println(err)
		fmt.Scanln()
		os.Exit(1)
	}

	// il file originario è codificato Windows-1252
	// viene convertito in UTF-8 altrimenti i caratteri speciali
	// non verranno riconosciuto
	file := charmap.Windows1252.NewDecoder().Reader(f)
	log.Println("Codifica caratteri da Windows-1252 a UTF-8")

	// Contenutore di tutte le tabelle trovate
	data := make(map[int][][]string)
	// Contatore per numerazione tabelle
	dataIndex := 0
	// Contenitore per la singola stringgga di dati
	dataRow := []string{}
	// Contatore per numero di celle contate in una riga
	// se non combacia con il numero di elementi nella riga
	// verrà aggiunta una cella vuota
	dataRowCount := 0

	// se settati ad uno vuol dire che ci si trova
	// all'interno del rispettivo elemento
	table := 0
	row := 0
	cell := 0

	z := html.NewTokenizer(file)

mainLoop:
	for {
		tt := z.Next()
		switch tt {

		// Esce dal loop in caso di file finito o in caso di errore
		case html.ErrorToken:
			log.Println(z.Err())
			break mainLoop

		// Aggiunge il testo alla riga nel caso sia all'interno di una tabella
		case html.TextToken:
			if cell > 0 {
				s := string(z.Text())
				dataRow = append(dataRow, s)
			}

		// Verifica che ci si trovi all'interno di una cella di una tabella
		case html.StartTagToken, html.EndTagToken:
			tn, _ := z.TagName()

			switch {
			// Verifico di errere all'interno di una tabella ed incremento il nome scheda
			case strings.Contains(string(tn), "table"):
				if tt == html.StartTagToken {
					table++
					dataIndex++
				} else {
					table--
				}

			// Verifico di essere all'interno di una riga
			// alla chiusura del tag TR aggiungo i dati all'array
			// e cancello quanto contenuto nella riga.
			case strings.Contains(string(tn), "tr") && table > 0:
				if tt == html.StartTagToken {
					row++
				} else {
					row--
					data[dataIndex] = append(data[dataIndex], dataRow)
					dataRow = []string{}
					dataRowCount = 0
				}

			// Verifico di essere all'interno di una cella titolo o normale
			// Ad ogni cella incremento il contatore, se chiudo una cella senza
			// aver inseriro dati nell'array inseriso una cella vuota
			// come sostituto
			case (strings.Contains(string(tn), "td") || strings.Contains(string(tn), "th")) && row > 0:
				if tt == html.StartTagToken {
					cell++
					dataRowCount++
				} else {
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

	// Per ogni tabella trovata creo una scheda nuova nell'excel
	// ad eccezione della scheda Sheet1 che già esiste
	for index := 1; index <= len(data); index++ {
		table := fmt.Sprint("Sheet", index)
		if index > 1 {
			xlsx.NewSheet(table)
		}

		// per ogni cella nell'array converto l'indice delle colonne
		// in lettere e vado a scrivere in quella cella
		for row := 0; row < len(data[index]); row++ {
			for col := 0; col < len(data[index][row]); col++ {
				ax := fmt.Sprint(excelize.ToAlphaString(col), row+1)
				xlsx.SetCellValue(table, ax, data[index][row][col])
			}
		}
	}

	// Il filename prende il nome file dal file html di origine
	filename := strings.Trim(path.Base(os.Args[1]), path.Ext(os.Args[1]))
	log.Println("Salvataggio file", filename)

	err = xlsx.SaveAs(fmt.Sprint(filename, ".xlsx"))
	if err != nil {
		log.Println(err)
		fmt.Scanln()
		os.Exit(1)
	}

	log.Println("Operazione completata, premere un tasto per chiudere...")
	fmt.Scanln()
}
