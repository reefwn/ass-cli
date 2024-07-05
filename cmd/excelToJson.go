package cmd

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var excelToJSONCmd = &cobra.Command{
	Use:   "excelToJson [inputPath] [outputPath] [sheetIndex]",
	Short: "Convert Excel to JSON",
	Args:  cobra.MinimumNArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		inputPath := args[0]
		var outputPath string
		sheetIndex := 0

		if len(args) > 1 {
			outputPath = args[1]
		} else {
			outputPath = strings.TrimSuffix(inputPath, filepath.Ext(inputPath)) + ".json"
		}

		if len(args) > 2 {
			index, err := strconv.Atoi(args[2])
			if err != nil {
				fmt.Println("Error: invalid sheetIndex", args[2])
				return
			}
			sheetIndex = index
		}

		err := excelToJSON(inputPath, outputPath, sheetIndex)
		if err != nil {
			fmt.Println("Error:", err)
		} else {
			fmt.Println("Successfully converted", inputPath, "to", outputPath)
		}
	},
}

func init() {
	rootCmd.AddCommand(excelToJSONCmd)
}

func excelToJSON(inputPath, outputPath string, sheetIndex int) error {
	// Open the Excel file
	f, err := excelize.OpenFile(inputPath)
	if err != nil {
		return err
	}

	// Get the sheet name by index
	sheetName := f.GetSheetName(sheetIndex)
	if sheetName == "" {
		return fmt.Errorf("sheetIndex %d is out of range", sheetIndex)
	}

	// Read all rows from the specified sheet
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return err
	}

	if len(rows) < 1 {
		return fmt.Errorf("the Excel file is empty")
	}

	// Create a slice of maps to hold the JSON data
	var data []map[string]interface{}
	headers := rows[0]

	// Convert each row to a map and append it to the data slice
	for _, row := range rows[1:] {
		record := make(map[string]interface{})
		for colIndex, cell := range row {
			if colIndex < len(headers) {
				record[headers[colIndex]] = cell
			}
		}
		data = append(data, record)
	}

	// Marshal the data to JSON
	jsonData, err := json.MarshalIndent(data, "", "  ")
	if err != nil {
		return err
	}

	// Write the JSON data to the output file
	if err := os.WriteFile(outputPath, jsonData, 0644); err != nil {
		return err
	}

	return nil
}
