package cmd

import (
	"fmt"
	"log"
	"net/url"

	"github.com/spf13/cobra"
)

var urlCmd = &cobra.Command{
	Use:   "url",
	Short: "URL encode/decode operations",
}

var urlEncodeCmd = &cobra.Command{
	Use:   "encode [string]",
	Short: "URL encode a string",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		fmt.Println("🔒", url.QueryEscape(args[0]))
	},
}

var urlDecodeCmd = &cobra.Command{
	Use:   "decode [encoded string]",
	Short: "URL decode a string",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		decoded, err := url.QueryUnescape(args[0])
		if err != nil {
			log.Fatalf("❌ %v", err)
		}
		fmt.Println("🔓", decoded)
	},
}

func init() {
	rootCmd.AddCommand(urlCmd)
	urlCmd.AddCommand(urlEncodeCmd)
	urlCmd.AddCommand(urlDecodeCmd)
}
