package cmd

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var inputPath string
var outputPath string
var sheetIndex int

var excelToJsonCmd = &cobra.Command{
	Use:   "excel2json",
	Short: "Convert an excel file to a json file",
	Long:  "Convert an excel file to a json file",
	Run: func(cmd *cobra.Command, args []string) {
		if inputPath == "" {
			log.Fatal("Input path must be provided")
		}

		if outputPath == "" {
			outputPath = defaultOutputPath(inputPath)
		}

		f, err := excelize.OpenFile(inputPath)
		if err != nil {
			log.Fatalf("Failed to open excel file: %v", err)
		}

		sheetName := f.GetSheetName(sheetIndex)
		if sheetName == "" {
			log.Fatalf("Sheet index %d does not exist", sheetIndex)
		}

		rows, err := f.GetRows(sheetName)
		if err != nil {
			log.Fatalf("Failed to get rows from sheet: %v", err)
		}

		if len(rows) < 1 {
			log.Fatal("No data found in sheet")
		}

		headers := rows[0]
		data := []map[string]string{}

		for _, row := range rows[1:] {
			if isEmptyRow(row) {
				continue
			}
			item := make(map[string]string)
			for i, cell := range row {
				if i < len(headers) {
					item[headers[i]] = cell
				}
			}
			data = append(data, item)
		}

		jsonData, err := json.Marshal(data)
		if err != nil {
			log.Fatalf("Failed to marshal data to JSON: %v", err)
		}

		err = os.WriteFile(outputPath, jsonData, 0644)
		if err != nil {
			log.Fatalf("Failed to write JSON to file: %v", err)
		}

		fmt.Printf("Excel file %s has been converted to JSON and saved to %s\n", inputPath, outputPath)
	},
}

func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if cell != "" {
			return false
		}
	}
	return true
}

func defaultOutputPath(inputPath string) string {
	ext := filepath.Ext(inputPath)
	name := inputPath[:len(inputPath)-len(ext)]
	return name + ".json"
}

func init() {
	rootCmd.AddCommand(excelToJsonCmd)

	excelToJsonCmd.Flags().StringVarP(&inputPath, "inputPath", "i", "", "Input Excel file path")
	excelToJsonCmd.Flags().StringVarP(&outputPath, "outputPath", "o", "", "Output JSON file path (default is same as input file name with .json extension)")
	excelToJsonCmd.Flags().IntVarP(&sheetIndex, "sheetIndex", "s", 0, "Sheet index (starting from 0)")
}
