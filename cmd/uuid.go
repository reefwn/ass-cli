package cmd

import (
	"fmt"
	"log"

	"github.com/atotto/clipboard"
	"github.com/google/uuid"
	"github.com/spf13/cobra"
)

var uuidCmd = &cobra.Command{
	Use:   "uuid [string]",
	Short: "Generate or validate a UUID",
	Long:  "Without args: generate a new UUID. With arg: validate the given string as UUID.",
	Run: func(cmd *cobra.Command, args []string) {
		if len(args) > 0 {
			if _, err := uuid.Parse(args[0]); err != nil {
				fmt.Println("❌ invalid")
			} else {
				fmt.Println("✅ valid")
			}
			return
		}

		id := uuid.New().String()
		if err := clipboard.WriteAll(id); err != nil {
			log.Fatalf("❌ %v", err)
		}
		fmt.Println("📋", id)
	},
}

func init() {
	rootCmd.AddCommand(uuidCmd)
}
