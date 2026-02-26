package cmd

import (
	"fmt"
	"log"
	"math/rand"

	"github.com/atotto/clipboard"
	"github.com/spf13/cobra"
)

var thaiIDCmd = &cobra.Command{
	Use:   "thai-id [id]",
	Short: "Generate or validate a Thai 13-digit national ID",
	Long:  "Without args: generate a random Thai ID. With arg: validate the given ID.",
	Run: func(cmd *cobra.Command, args []string) {
		if len(args) > 0 {
			if validateThaiID(args[0]) {
				fmt.Println("✅ valid")
			} else {
				fmt.Println("❌ invalid")
			}
			return
		}

		id := generateThaiID()
		if err := clipboard.WriteAll(id); err != nil {
			log.Fatalf("❌ %v", err)
		}
		fmt.Println("📋", id)
	},
}

func validateThaiID(id string) bool {
	if len(id) != 13 {
		return false
	}
	sum := 0
	for i := 0; i < 12; i++ {
		if id[i] < '0' || id[i] > '9' {
			return false
		}
		sum += int(id[i]-'0') * (13 - i)
	}
	return id[12]-'0' == byte((11-(sum%11))%10)
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

func init() {
	rootCmd.AddCommand(thaiIDCmd)
}
