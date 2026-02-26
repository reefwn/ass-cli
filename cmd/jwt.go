package cmd

import (
	"encoding/json"
	"fmt"
	"log"
	"strings"

	"github.com/golang-jwt/jwt/v5"
	"github.com/spf13/cobra"
)

var jwtCmd = &cobra.Command{
	Use:   "jwt",
	Short: "JWT decode/sign operations",
}

var jwtDecodeCmd = &cobra.Command{
	Use:   "decode [token]",
	Short: "Decode a JWT token",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		tokenStr := args[0]
		if strings.HasPrefix(strings.ToLower(tokenStr), "bearer ") {
			tokenStr = strings.TrimSpace(tokenStr[7:])
		}

		token, _, err := new(jwt.Parser).ParseUnverified(tokenStr, jwt.MapClaims{})
		if err != nil {
			log.Fatalf("❌ %v", err)
		}

		claims, ok := token.Claims.(jwt.MapClaims)
		if !ok {
			log.Fatal("❌ could not parse claims")
		}

		prettyJSON, err := json.MarshalIndent(claims, "", "  ")
		if err != nil {
			log.Fatalf("❌ %v", err)
		}
		fmt.Println(string(prettyJSON))
	},
}

var jwtSignCmd = &cobra.Command{
	Use:   "sign [json] [secret]",
	Short: "Sign JSON payload to JWT",
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		var claims jwt.MapClaims
		if err := json.Unmarshal([]byte(args[0]), &claims); err != nil {
			log.Fatalf("❌ %v", err)
		}

		token := jwt.NewWithClaims(jwt.SigningMethodHS256, claims)
		signed, err := token.SignedString([]byte(args[1]))
		if err != nil {
			log.Fatalf("❌ %v", err)
		}
		fmt.Println("🔑", signed)
	},
}

func init() {
	rootCmd.AddCommand(jwtCmd)
	jwtCmd.AddCommand(jwtDecodeCmd)
	jwtCmd.AddCommand(jwtSignCmd)
}
