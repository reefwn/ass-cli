package cmd

import (
	"fmt"
	"log"
	"strconv"
	"time"

	"github.com/spf13/cobra"
)

var unixToUtcCmd = &cobra.Command{
	Use:   "unixToUtc [timestamp]",
	Short: "Convert unix timestamp to UTC datetime",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		ts, err := strconv.ParseInt(args[0], 10, 64)
		if err != nil {
			log.Fatalf("Invalid timestamp: %v", err)
		}

		// auto-detect milliseconds (13+ digits)
		var t time.Time
		if ts > 9999999999 {
			t = time.UnixMilli(ts).UTC()
		} else {
			t = time.Unix(ts, 0).UTC()
		}

		result := t.Format(time.RFC3339)

		fmt.Println(result)
	},
}

func init() {
	rootCmd.AddCommand(unixToUtcCmd)
}
