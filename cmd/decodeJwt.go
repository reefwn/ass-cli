package cmd

import (
	"encoding/json"
	"fmt"
	"log"
	"strings"

	"github.com/golang-jwt/jwt/v5"
	"github.com/spf13/cobra"
)

var decodeJwtCmd = &cobra.Command{
	Use:   "decodeJwt [token]",
	Short: "Decode a JWT token",
	Run: func(cmd *cobra.Command, args []string) {
		tokenStr := args[0]

		if strings.HasPrefix(strings.ToLower(tokenStr), "bearer ") {
			tokenStr = strings.TrimSpace(tokenStr[7:])
		}

		token, _, err := new(jwt.Parser).ParseUnverified(tokenStr, jwt.MapClaims{})
		if err != nil {
			log.Fatalf("Failed to decode token: %v", err)
		}

		if claims, ok := token.Claims.(jwt.MapClaims); ok {
			prettyJSON, err := json.MarshalIndent(claims, "", "  ")
			if err != nil {
				log.Fatalf("Failed to format claims: %v", err)
			}
			fmt.Println(string(prettyJSON))
		} else {
			log.Println("Could not parse claims")
		}
	},
}

func init() {
	rootCmd.AddCommand(decodeJwtCmd)
}
