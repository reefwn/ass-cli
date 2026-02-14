package cmd

import (
	"fmt"
	"log"
	"math/rand"

	"github.com/atotto/clipboard"
	"github.com/spf13/cobra"
)

func validateThaiID(id string) bool {
	if len(id) != 13 {
		return false
	}
	sum := 0
	for i := 0; i < 12; i++ {
		d := int(id[i] - '0')
		if id[i] < '0' || id[i] > '9' {
			return false
		}
		sum += d * (13 - i)
	}
	check := (11 - (sum % 11)) % 10
	return id[12]-'0' == byte(check)
}

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
	Short: "Generate or validate a Thai 13-digit national ID",
	Run: func(cmd *cobra.Command, args []string) {
		if len(args) > 0 {
			if validateThaiID(args[0]) {
				fmt.Println("Valid Thai ID")
			} else {
				fmt.Println("Invalid Thai ID")
			}
			return
		}

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
