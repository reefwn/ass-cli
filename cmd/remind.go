package cmd

import (
	"fmt"
	"os"
	"os/exec"
	"time"

	"github.com/spf13/cobra"
)

var remindCmd = &cobra.Command{
	Use:   "remind [duration] [message]",
	Short: "Set a reminder notification",
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		delayTime, err := time.ParseDuration(args[0])
		if err != nil {
			fmt.Println("❌ invalid duration, use e.g. '10s', '5m', '1h'")
			os.Exit(1)
		}

		message := args[1]
		totalSeconds := int(delayTime.Seconds())
		fmt.Printf("⏳ %s\n", delayTime)

		for i := totalSeconds; i > 0; i-- {
			fmt.Printf("\r⏳ %ds", i)
			time.Sleep(1 * time.Second)
		}
		fmt.Println()

		scriptCmd := exec.Command("osascript", "-e", fmt.Sprintf(`display notification "%s" with title "Notification"`, message))
		if err := scriptCmd.Run(); err != nil {
			fmt.Println("❌", err)
			os.Exit(1)
		}
		fmt.Println("🔔", message)
	},
}

func init() {
	rootCmd.AddCommand(remindCmd)
}
