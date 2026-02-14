package cmd

import (
	"fmt"
	"log"
	"math/rand"

	"github.com/atotto/clipboard"
	"github.com/spf13/cobra"
)

func generateThaiID() string {
	digits := make([]int, 13)
	for i := 0; i < 12; i++ {
		digits[i] = rand.Intn(10)
	}
	sum := 0
	for i := 0; i < 12; i++ {
		sum += digits[i] * (13 - i)
	}
	digits[12] = (11 - (sum % 11)) % 10

	result := make([]byte, 13)
	for i, d := range digits {
		result[i] = byte('0' + d)
	}
	return string(result)
}

var thaiIDCmd = &cobra.Command{
	Use:   "thaiID",
	Short: "Generate a Thai 13-digit national ID",
	Run: func(cmd *cobra.Command, args []string) {
		id := generateThaiID()

		if err := clipboard.WriteAll(id); err != nil {
			log.Fatalf("Failed to copy to clipboard: %v", err)
		}

		fmt.Println("copied to clipboard:", id)
	},
}

func init() {
	rootCmd.AddCommand(thaiIDCmd)
}
