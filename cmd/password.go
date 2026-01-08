package cmd

import (
	"crypto/rand"
	"fmt"
	"log"
	"strconv"

	"github.com/atotto/clipboard"
	"github.com/spf13/cobra"
)

var passwordCmd = &cobra.Command{
	Use:   "password [length]",
	Short: "Generate a random password",
	Long:  "Generate a random password of specified length. Defaults to 12 characters using letters and numbers. Use --special to include special characters.",
	Run: func(cmd *cobra.Command, args []string) {
		length := 12
		if len(args) > 0 {
			var err error
			length, err = strconv.Atoi(args[0])
			if err != nil {
				fmt.Println("Invalid length. Please provide a valid number.")
				return
			}
			if length <= 0 {
				fmt.Println("Length must be greater than 0.")
				return
			}
		}

		special, _ := cmd.Flags().GetBool("special")
		charset := "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
		if special {
			charset += "!@#$%^&*()_+-=[]{}|;:,.<>?"
		}

		password := generatePassword(length, charset)

		if err := clipboard.WriteAll(password); err != nil {
			log.Fatalf("Failed to copy to clipboard: %v", err)
		}

		fmt.Println("Generated password copied to clipboard:", password)
	},
}

func generatePassword(length int, charset string) string {
	password := make([]byte, length)
	charsetLen := len(charset)

	randomBytes := make([]byte, length)
	_, err := rand.Read(randomBytes)
	if err != nil {
		log.Fatalf("Failed to generate random bytes: %v", err)
	}

	for i := 0; i < length; i++ {
		password[i] = charset[randomBytes[i]%byte(charsetLen)]
	}

	return string(password)
}

func init() {
	rootCmd.AddCommand(passwordCmd)
	passwordCmd.Flags().BoolP("special", "s", false, "Include special characters in the password")
}
