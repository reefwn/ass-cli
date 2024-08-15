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

var imgToBase64Cmd = &cobra.Command{
	Use:   "imgToBase64 [image path]",
	Short: "Convert given image path to base64 string with MIME type",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		imagePath := args[0]
		base64String, mimeType, err := imageToBase64(imagePath)
		if err != nil {
			log.Fatalf("Error converting image to base64: %v", err)
		}

		// Copy the Base64 string to the clipboard
		if err := clipboard.WriteAll(base64String); err != nil {
			log.Fatalf("Failed to copy to clipboard: %v", err)
		}

		fmt.Printf("Base64 copied to clipboard\nFile Path: %s\nMIME Type: %s\n", imagePath, mimeType)
	},
}

func init() {
	rootCmd.AddCommand(imgToBase64Cmd)
}

func imageToBase64(imagePath string) (string, string, error) {
	// Read the image file
	imageData, err := os.ReadFile(imagePath)
	if err != nil {
		return "", "", err
	}

	// Get the MIME type
	ext := filepath.Ext(imagePath)
	mimeType := mime.TypeByExtension(ext)
	if mimeType == "" {
		return "", "", fmt.Errorf("Could not determine MIME type for file extension: %s", ext)
	}

	// Encode the image data to Base64
	base64String := base64.StdEncoding.EncodeToString(imageData)

	return base64String, mimeType, nil
}
