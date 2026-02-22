package cmd

import (
	"encoding/json"
	"fmt"
	"log"

	"github.com/golang-jwt/jwt/v5"
	"github.com/spf13/cobra"
)

var signJwtCmd = &cobra.Command{
	Use:   "signJwt [json] [secret]",
	Short: "Sign JSON payload to JWT",
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		var claims jwt.MapClaims
		if err := json.Unmarshal([]byte(args[0]), &claims); err != nil {
			log.Fatalf("Invalid JSON: %v", err)
		}

		token := jwt.NewWithClaims(jwt.SigningMethodHS256, claims)
		signed, err := token.SignedString([]byte(args[1]))
		if err != nil {
			log.Fatalf("Failed to sign token: %v", err)
		}

		fmt.Println(signed)
	},
}

func init() {
	rootCmd.AddCommand(signJwtCmd)
}
