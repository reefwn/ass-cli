package cmd

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"
)

var csvCmd = &cobra.Command{
	Use:   "csv",
	Short: "CSV conversion operations",
}

var csvToJsonCmd = &cobra.Command{
	Use:   "to-json [inputPath] [outputPath]",
	Short: "Convert CSV to JSON",
	Args:  cobra.MinimumNArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		inputPath := args[0]
		outputPath := strings.TrimSuffix(inputPath, filepath.Ext(inputPath)) + ".json"
		if len(args) > 1 {
			outputPath = args[1]
		}

		csvFile, err := os.Open(inputPath)
		if err != nil {
			fmt.Println("❌", err)
			return
		}
		defer csvFile.Close()

		records, err := csv.NewReader(csvFile).ReadAll()
		if err != nil {
			fmt.Println("❌", err)
			return
		}
		if len(records) < 2 {
			fmt.Println("❌ need at least one data row")
			return
		}

		headers := records[0]
		var jsonData []map[string]string
		for _, row := range records[1:] {
			entry := make(map[string]string)
			for i, value := range row {
				entry[headers[i]] = value
			}
			jsonData = append(jsonData, entry)
		}

		jsonBytes, err := json.MarshalIndent(jsonData, "", "  ")
		if err != nil {
			fmt.Println("❌", err)
			return
		}
		if err := os.WriteFile(outputPath, jsonBytes, 0644); err != nil {
			fmt.Println("❌", err)
			return
		}
		fmt.Println("✅", outputPath)
	},
}

func init() {
	rootCmd.AddCommand(csvCmd)
	csvCmd.AddCommand(csvToJsonCmd)
}
