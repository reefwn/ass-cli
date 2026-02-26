package cmd

import (
	"crypto/rand"
	"fmt"
	"log"
	"strconv"

	"github.com/atotto/clipboard"
	"github.com/spf13/cobra"
)

var passwordSpecial bool
var passwordLength int

var passwordCmd = &cobra.Command{
	Use:   "password",
	Short: "Generate a random password",
	Run: func(cmd *cobra.Command, args []string) {
		charset := "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
		if passwordSpecial {
			charset += "!@#$%^&*()_+-=[]{}|;:,.<>?"
		}

		password := make([]byte, passwordLength)
		randomBytes := make([]byte, passwordLength)
		if _, err := rand.Read(randomBytes); err != nil {
			log.Fatalf("❌ %v", err)
		}
		for i := range password {
			password[i] = charset[randomBytes[i]%byte(len(charset))]
		}

		result := string(password)
		if err := clipboard.WriteAll(result); err != nil {
			log.Fatalf("❌ %v", err)
		}
		fmt.Println("📋", result)
	},
}

func init() {
	rootCmd.AddCommand(passwordCmd)
	passwordCmd.Flags().IntVarP(&passwordLength, "length", "l", 12, "Password length")
	passwordCmd.Flags().BoolVarP(&passwordSpecial, "special", "s", false, "Include special characters")

	// Keep backward compat: allow `ass password 20` positional arg
	passwordCmd.Args = cobra.MaximumNArgs(1)
	originalRun := passwordCmd.Run
	passwordCmd.Run = func(cmd *cobra.Command, args []string) {
		if len(args) > 0 && !cmd.Flags().Changed("length") {
			if l, err := strconv.Atoi(args[0]); err == nil && l > 0 {
				passwordLength = l
			} else {
				fmt.Println("❌ invalid length")
				return
			}
		}
		originalRun(cmd, args)
	}
}
