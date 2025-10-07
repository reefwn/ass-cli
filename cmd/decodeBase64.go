package cmd

import (
	"encoding/base64"
	"fmt"

	"github.com/spf13/cobra"
)

var decodeBase64Cmd = &cobra.Command{
	Use:   "decodeBase64 [base64]",
	Short: "Decode a Base64 string",
	Run: func(cmd *cobra.Command, args []string) {
		if len(args) == 0 {
			fmt.Println("Please provide a Base64 string")
			return
		}

		input := args[0]

		decoded, err := DecodeBase64(input)
		if err != nil {
			fmt.Println("Error decoding Base64:", err)
			return
		}

		fmt.Println("Decoded string:", decoded)
	},
}

func init() {
	rootCmd.AddCommand(decodeBase64Cmd)
}

func DecodeBase64(encoded string) (string, error) {
	data, err := base64.StdEncoding.DecodeString(encoded)
	if err != nil {
		return "", err
	}
	return string(data), nil
}
