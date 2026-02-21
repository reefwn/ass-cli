package cmd

import (
	"encoding/base64"
	"fmt"

	"github.com/spf13/cobra"
)

var encodeBase64Cmd = &cobra.Command{
	Use:   "encodeBase64 [string]",
	Short: "Encode a string to Base64",
	Run: func(cmd *cobra.Command, args []string) {
		if len(args) == 0 {
			fmt.Println("Please provide a string to encode")
			return
		}
		fmt.Println("Encoded string:", base64.StdEncoding.EncodeToString([]byte(args[0])))
	},
}

func init() {
	rootCmd.AddCommand(encodeBase64Cmd)
}
