package cmd

import (
	"encoding/base64"
	"fmt"
	"log"
	"mime"
	"os"
	"path/filepath"

	"github.com/atotto/clipboard"
	"github.com/spf13/cobra"
)

var base64Cmd = &cobra.Command{
	Use:   "base64",
	Short: "Base64 encode/decode operations",
}

var base64EncodeCmd = &cobra.Command{
	Use:   "encode [string]",
	Short: "Encode a string to Base64",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		fmt.Println("🔒", base64.StdEncoding.EncodeToString([]byte(args[0])))
	},
}

var base64DecodeCmd = &cobra.Command{
	Use:   "decode [base64]",
	Short: "Decode a Base64 string",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		data, err := base64.StdEncoding.DecodeString(args[0])
		if err != nil {
			log.Fatalf("❌ %v", err)
		}
		fmt.Println("🔓", string(data))
	},
}

var base64ImgCmd = &cobra.Command{
	Use:   "img [image path]",
	Short: "Convert image to Base64 with MIME type",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		imageData, err := os.ReadFile(args[0])
		if err != nil {
			log.Fatalf("❌ %v", err)
		}

		mimeType := mime.TypeByExtension(filepath.Ext(args[0]))
		if mimeType == "" {
			log.Fatalf("❌ unknown MIME type: %s", filepath.Ext(args[0]))
		}

		encoded := base64.StdEncoding.EncodeToString(imageData)
		if err := clipboard.WriteAll(encoded); err != nil {
			log.Fatalf("❌ %v", err)
		}
		fmt.Printf("📋 %s (%s)\n", args[0], mimeType)
	},
}

func init() {
	rootCmd.AddCommand(base64Cmd)
	base64Cmd.AddCommand(base64EncodeCmd)
	base64Cmd.AddCommand(base64DecodeCmd)
	base64Cmd.AddCommand(base64ImgCmd)
}
