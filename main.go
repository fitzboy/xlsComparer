package main

import (
	"flag"
	"log"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

var older = flag.String("older", "", "path to older file")
var newer = flag.String("newer", "", "path to newer file")
var output = flag.String("output", "", "path to output file")

func main() {
	flag.Parse()
	if *older == "" || *newer == "" || *output == "" {
		log.Fatal("one of the filenames is missing, use -h to see options")
	}

	olderFile, err := xlsx.OpenFile(*older)
	if err != nil {
		log.Fatalf("unable to read %s because %v", *older, err)
	}
	valuesSeen := make(map[string]bool)
	for _, sheet := range olderFile.Sheets {
		for _, row := range sheet.Rows {
			var sb strings.Builder
			for _, cell := range row.Cells {
				sb.WriteString(cell.String())
			}
			text := sb.String()
			if text != "" {
				valuesSeen[text] = true
			}
		}
	}

	outputFile := xlsx.NewFile()
	newerFile, err := xlsx.OpenFile(*newer)
	if err != nil {
		log.Fatalf("unable to read %s because %v", *newer, err)
	}

	for _, newerSheet := range newerFile.Sheets {
		outputSheet, err := outputFile.AddSheet(newerSheet.Name)
		if err != nil {
			log.Fatalf("unable to add sheet %s", newerSheet.Name)
		}
		for _, newerRow := range newerSheet.Rows {
			outputRow := outputSheet.AddRow()
			wasSeen := false
			var sb strings.Builder
			for _, newerCell := range newerRow.Cells {
				sb.WriteString(newerCell.String())
			}
			text := sb.String()
			if _, ok := valuesSeen[text]; ok {
				wasSeen = true
				for _, newerCell := range newerRow.Cells {
					s := xlsx.NewStyle()
					s.Fill = *xlsx.FillGreen
					newerCell.SetStyle(s)
				}
			}

			for _, newerCell := range newerRow.Cells {
				outputCell := outputRow.AddCell()
				text := newerCell.String()
				outputCell.Value = text
				s := xlsx.NewStyle()
				if wasSeen {
					s.Fill = *xlsx.FillGreen
					//					fmt.Printf("green: %s\n", text)
				} else {
					s.Fill = *xlsx.FillRed
					//					fmt.Printf("red: %s\n", text)
				}
				s.ApplyFill = true
				s.Font.Size = 8
				s.Font.Name = "Arial"
				//				s.Alignment.ShrinkToFit = true
				//				newerStyle := newerCell.GetStyle()
				//				s.Border = newerStyle.Border
				//				s.ApplyBorder = newerStyle.ApplyBorder
				//				s.Font = newerStyle.Font
				//				s.ApplyFont = newerStyle.ApplyFont
				//				s.Alignment = newerStyle.Alignment
				//				s.ApplyAlignment = newerStyle.ApplyAlignment
				outputCell.SetStyle(s)
			}
		}
	}

	os.Remove(*output)
	err = newerFile.Save(*output)
	if err != nil {
		log.Fatalf("unable to save output file %s because %v", *output, err)
	}
}
