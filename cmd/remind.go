package cmd

import (
	"fmt"
	"os"
	"os/exec"
	"time"

	"github.com/spf13/cobra"
)

var remindCmd = &cobra.Command{
	Use:   "remind",
	Short: "Notify me",
	Run: func(cmd *cobra.Command, args []string) {
		// Parse the delay time with a unit (e.g., "10s", "5m", "1h")
		delayTime, err := time.ParseDuration(args[0])
		if err != nil {
			fmt.Println("Invalid delay time. Please provide a valid duration (e.g., '10s', '5m', '1h').")
			os.Exit(1)
		}

		// Get the notification message from the second argument
		message := args[1]

		// Start countdown
		totalSeconds := int(delayTime.Seconds())
		fmt.Printf("Countdown started for %s...\n", delayTime)

		for i := totalSeconds; i > 0; i-- {
			fmt.Printf("\rTime remaining: %d seconds", i)
			time.Sleep(1 * time.Second)
		}
		fmt.Println("\nCountdown complete.")

		// Send the macOS notification
		scriptCmd := exec.Command("osascript", "-e", fmt.Sprintf(`display notification "%s" with title "Notification"`, message))
		err = scriptCmd.Run()
		if err != nil {
			fmt.Println("Error sending notification:", err)
			os.Exit(1)
		}

		fmt.Println("Notification sent successfully!")
	},
}

func init() {
	rootCmd.AddCommand(remindCmd)
}
