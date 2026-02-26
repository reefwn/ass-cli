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

var excelCmd = &cobra.Command{
	Use:   "excel",
	Short: "Excel conversion operations",
}

var excelToJsonCmd = &cobra.Command{
	Use:   "to-json [inputPath] [outputPath] [sheetIndex]",
	Short: "Convert Excel to JSON",
	Args:  cobra.MinimumNArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		inputPath := args[0]
		outputPath := strings.TrimSuffix(inputPath, filepath.Ext(inputPath)) + ".json"
		sheetIndex := 0

		if len(args) > 1 {
			outputPath = args[1]
		}
		if len(args) > 2 {
			idx, err := strconv.Atoi(args[2])
			if err != nil {
				fmt.Println("❌ invalid sheetIndex", args[2])
				return
			}
			sheetIndex = idx
		}

		f, err := excelize.OpenFile(inputPath)
		if err != nil {
			fmt.Println("❌", err)
			return
		}

		sheetName := f.GetSheetName(sheetIndex)
		if sheetName == "" {
			fmt.Printf("❌ sheetIndex %d out of range\n", sheetIndex)
			return
		}

		rows, err := f.GetRows(sheetName)
		if err != nil {
			fmt.Println("❌", err)
			return
		}
		if len(rows) < 1 {
			fmt.Println("❌ empty file")
			return
		}

		headers := rows[0]
		var data []map[string]interface{}
		for _, row := range rows[1:] {
			record := make(map[string]interface{})
			for colIndex, cell := range row {
				if colIndex < len(headers) {
					record[headers[colIndex]] = cell
				}
			}
			data = append(data, record)
		}

		jsonData, err := json.MarshalIndent(data, "", "  ")
		if err != nil {
			fmt.Println("❌", err)
			return
		}
		if err := os.WriteFile(outputPath, jsonData, 0644); err != nil {
			fmt.Println("❌", err)
			return
		}
		fmt.Println("✅", outputPath)
	},
}

var excelFromJsonCmd = &cobra.Command{
	Use:   "from-json [inputPath] [outputPath] [sheetName]",
	Short: "Convert JSON to Excel",
	Args:  cobra.MinimumNArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		inputPath := args[0]
		outputPath := strings.TrimSuffix(inputPath, filepath.Ext(inputPath)) + ".xlsx"
		sheetName := "Sheet1"

		if len(args) > 1 {
			outputPath = args[1]
		}
		if len(args) > 2 {
			sheetName = args[2]
		}

		jsonData, err := os.ReadFile(inputPath)
		if err != nil {
			fmt.Println("❌", err)
			return
		}

		var data []map[string]interface{}
		if err := json.Unmarshal(jsonData, &data); err != nil {
			fmt.Println("❌", err)
			return
		}

		f := excelize.NewFile()
		f.NewSheet(sheetName)

		if len(data) > 0 {
			headers := make([]string, 0, len(data[0]))
			for key := range data[0] {
				headers = append(headers, key)
			}
			for colIndex, header := range headers {
				cell, _ := excelize.CoordinatesToCellName(colIndex+1, 1)
				f.SetCellValue(sheetName, cell, header)
			}
			for rowIndex, record := range data {
				for colIndex, header := range headers {
					cell, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+2)
					f.SetCellValue(sheetName, cell, record[header])
				}
			}
		}

		if err := f.SaveAs(outputPath); err != nil {
			fmt.Println("❌", err)
			return
		}
		fmt.Println("✅", outputPath)
	},
}

func init() {
	rootCmd.AddCommand(excelCmd)
	excelCmd.AddCommand(excelToJsonCmd)
	excelCmd.AddCommand(excelFromJsonCmd)
}
