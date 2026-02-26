package cmd

import (
	"fmt"
	"log"
	"strconv"
	"time"

	"github.com/atotto/clipboard"
	"github.com/spf13/cobra"
)

var (
	unixMs      bool
	unixFromUtc string
	unixToUtc   string
)

var unixCmd = &cobra.Command{
	Use:   "unix",
	Short: "Get current unix timestamp, or convert between unix and UTC",
	Long: `Without flags: get current unix timestamp.
  --from-utc <datetime>  Convert UTC (RFC3339) to unix timestamp
  --to-utc <timestamp>   Convert unix timestamp to UTC datetime`,
	Run: func(cmd *cobra.Command, args []string) {
		if unixFromUtc != "" {
			t, err := time.Parse(time.RFC3339, unixFromUtc)
			if err != nil {
				log.Fatalf("❌ %v", err)
			}
			fmt.Println("🕐", t.Unix())
			return
		}

		if unixToUtc != "" {
			ts, err := strconv.ParseInt(unixToUtc, 10, 64)
			if err != nil {
				log.Fatalf("❌ %v", err)
			}
			var t time.Time
			if ts > 9999999999 {
				t = time.UnixMilli(ts).UTC()
			} else {
				t = time.Unix(ts, 0).UTC()
			}
			fmt.Println("🕐", t.Format(time.RFC3339))
			return
		}

		now := time.Now()
		var ts string
		if unixMs {
			ts = strconv.FormatInt(now.UnixMilli(), 10)
		} else {
			ts = strconv.FormatInt(now.Unix(), 10)
		}
		if err := clipboard.WriteAll(ts); err != nil {
			log.Fatalf("❌ %v", err)
		}
		fmt.Println("📋", ts)
	},
}

func init() {
	rootCmd.AddCommand(unixCmd)
	unixCmd.Flags().BoolVarP(&unixMs, "ms", "m", false, "Get timestamp in milliseconds")
	unixCmd.Flags().StringVar(&unixFromUtc, "from-utc", "", "Convert UTC datetime (RFC3339) to unix")
	unixCmd.Flags().StringVar(&unixToUtc, "to-utc", "", "Convert unix timestamp to UTC datetime")
}
