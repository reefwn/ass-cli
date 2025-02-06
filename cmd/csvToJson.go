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

var csvToJsonCmd = &cobra.Command{
	Use:   "csvToJson [inputPath] [outputPath]",
	Short: "Convert CSV to JSON",
	Run: func(cmd *cobra.Command, args []string) {
		inputPath := args[0]
		var outputPath string

		if len(args) > 1 {
			outputPath = args[1]
		} else {
			outputPath = strings.TrimSuffix(inputPath, filepath.Ext(inputPath)) + ".json"
		}

		err := csvToJSON(inputPath, outputPath)
		if err != nil {
			fmt.Println("Error:", err)
		} else {
			fmt.Println("Successfully converted", inputPath, "to", outputPath)
		}
	},
}

func init() {
	rootCmd.AddCommand(csvToJsonCmd)
}

func csvToJSON(inputPath, outputPath string) error {
	csvFile, err := os.Open(inputPath)
	if err != nil {
		fmt.Println("Error opening CSV file:", err)
		return err
	}
	defer csvFile.Close()

	reader := csv.NewReader(csvFile)
	records, err := reader.ReadAll()
	if err != nil {
		fmt.Println("Error reading CSV file:", err)
		return err
	}

	if len(records) < 2 {
		fmt.Println("CSV file must have at least one data row.")
		return err
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
		fmt.Println("Error converting to JSON:", err)
		return err
	}

	err = os.WriteFile(outputPath, jsonBytes, 0644)
	if err != nil {
		fmt.Println("Error writing JSON file:", err)
		return err
	}

	return nil
}
