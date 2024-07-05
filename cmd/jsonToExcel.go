package cmd

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var convertCmd = &cobra.Command{
	Use:   "jsonToExcel [inputPath] [outputPath] [sheetName]",
	Short: "Convert JSON to Excel",
	Args:  cobra.MinimumNArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		inputPath := args[0]
		var outputPath string
		sheetName := "Sheet1"

		if len(args) > 1 {
			outputPath = args[1]
		} else {
			outputPath = strings.TrimSuffix(inputPath, filepath.Ext(inputPath)) + ".xlsx"
		}

		if len(args) > 2 {
			sheetName = args[2]
		}

		err := jsonToExcel(inputPath, outputPath, sheetName)
		if err != nil {
			fmt.Println("Error:", err)
		} else {
			fmt.Println("Successfully converted", inputPath, "to", outputPath)
		}
	},
}

func init() {
	rootCmd.AddCommand(convertCmd)
}

func jsonToExcel(inputPath, outputPath, sheetName string) error {
	// Read JSON file
	jsonData, err := os.ReadFile(inputPath)
	if err != nil {
		return err
	}

	// Unmarshal JSON data
	var data []map[string]interface{}
	err = json.Unmarshal(jsonData, &data)
	if err != nil {
		return err
	}

	// Create a new Excel file
	f := excelize.NewFile()
	f.NewSheet(sheetName)

	if len(data) > 0 {
		// Write header
		headers := make([]string, 0, len(data[0]))
		for key := range data[0] {
			headers = append(headers, key)
		}

		for colIndex, header := range headers {
			cell, _ := excelize.CoordinatesToCellName(colIndex+1, 1)
			f.SetCellValue(sheetName, cell, header)
		}

		// Write rows
		for rowIndex, record := range data {
			for colIndex, header := range headers {
				cell, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+2)
				f.SetCellValue(sheetName, cell, record[header])
			}
		}
	}

	// Save Excel file
	if err := f.SaveAs(outputPath); err != nil {
		return err
	}

	return nil
}
