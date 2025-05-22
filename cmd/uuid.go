package cmd

import (
	"fmt"
	"log"

	"github.com/atotto/clipboard"
	"github.com/google/uuid"
	"github.com/spf13/cobra"
)

var uuidCmd = &cobra.Command{
	Use:   "uuid",
	Short: "Generate a UUID",
	Run: func(cmd *cobra.Command, args []string) {
		id := uuid.New()
		idStr := id.String()

		if err := clipboard.WriteAll(idStr); err != nil {
			log.Fatalf("Failed to copy to clipboard: %v", err)
		}

		fmt.Println("copied to clipboard:", idStr)
	},
}

func init() {
	rootCmd.AddCommand(uuidCmd)
}
