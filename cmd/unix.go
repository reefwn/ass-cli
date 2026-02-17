package cmd

import (
	"fmt"
	"log"
	"strconv"
	"time"

	"github.com/atotto/clipboard"
	"github.com/spf13/cobra"
)

var unixMs bool

var unixCmd = &cobra.Command{
	Use:   "unix",
	Short: "Get current unix timestamp",
	Run: func(cmd *cobra.Command, args []string) {
		now := time.Now()
		var ts string
		if unixMs {
			ts = strconv.FormatInt(now.UnixMilli(), 10)
		} else {
			ts = strconv.FormatInt(now.Unix(), 10)
		}

		if err := clipboard.WriteAll(ts); err != nil {
			log.Fatalf("Failed to copy to clipboard: %v", err)
		}

		fmt.Println("copied to clipboard:", ts)
	},
}

func init() {
	rootCmd.AddCommand(unixCmd)
	unixCmd.Flags().BoolVarP(&unixMs, "ms", "m", false, "Get timestamp in milliseconds")
}
