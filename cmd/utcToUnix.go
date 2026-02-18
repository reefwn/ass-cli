package cmd

import (
	"fmt"
	"log"
	"time"

	"github.com/spf13/cobra"
)

var utcToUnixCmd = &cobra.Command{
	Use:   "utcToUnix [datetime]",
	Short: "Convert UTC datetime to unix timestamp",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		t, err := time.Parse(time.RFC3339, args[0])
		if err != nil {
			log.Fatalf("Invalid datetime (expected RFC3339 format): %v", err)
		}
		fmt.Println(t.Unix())
	},
}

func init() {
	rootCmd.AddCommand(utcToUnixCmd)
}
