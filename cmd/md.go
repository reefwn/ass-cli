package cmd

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/gomarkdown/markdown"
	"github.com/gomarkdown/markdown/parser"
	"github.com/jung-kurt/gofpdf"
	"github.com/spf13/cobra"
)

var mdCmd = &cobra.Command{
	Use:   "md",
	Short: "Markdown operations",
}

var mdToPdfCmd = &cobra.Command{
	Use:   "to-pdf [file]",
	Short: "Convert markdown file to PDF",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		inputPath := args[0]

		if _, err := os.Stat(inputPath); os.IsNotExist(err) {
			fmt.Printf("❌ file not found: %s\n", inputPath)
			return
		}

		mdContent, err := os.ReadFile(inputPath)
		if err != nil {
			fmt.Printf("❌ %v\n", err)
			return
		}

		ext := filepath.Ext(inputPath)
		outputPath := strings.TrimSuffix(inputPath, ext) + ".pdf"

		extensions := parser.CommonExtensions | parser.AutoHeadingIDs
		p := parser.NewWithExtensions(extensions)
		htmlContent := markdown.ToHTML(mdContent, p, nil)

		pdf := gofpdf.New("P", "mm", "A4", "")
		pdf.SetMargins(15, 15, 15)
		pdf.AddPage()
		pdf.SetFont("Arial", "", 12)

		renderHTMLToPDF(pdf, string(htmlContent))

		if err := pdf.OutputFileAndClose(outputPath); err != nil {
			fmt.Printf("❌ %v\n", err)
			return
		}
		fmt.Println("✅", outputPath)
	},
}

func renderHTMLToPDF(pdf *gofpdf.Fpdf, htmlContent string) {
	htmlContent = regexp.MustCompile(`<!--.*?-->`).ReplaceAllString(htmlContent, "")

	tagRe := regexp.MustCompile(`<(/?)([a-zA-Z0-9]+)[^>]*>`)

	fontStyle := ""
	fontSize := 12.0
	listDepth := 0
	baseLeftMargin := 15.0
	justStartedListItem := false

	inTable := false
	inTableHeader := false
	tableRow := []tableCell{}
	tableColWidths := []float64{}
	cellText := ""
	cellStyle := ""
	inCell := false

	lastEnd := 0
	matches := tagRe.FindAllStringSubmatchIndex(htmlContent, -1)

	for _, match := range matches {
		if match[0] > lastEnd {
			text := strings.TrimSpace(htmlContent[lastEnd:match[0]])
			if text != "" {
				text = decodeHTMLEntities(text)
				if inCell {
					cellText += text
				} else {
					pdf.Write(5, text+" ")
					justStartedListItem = false
				}
			}
		}
		lastEnd = match[1]

		isClosing := htmlContent[match[2]:match[3]] == "/"
		tagName := strings.ToLower(htmlContent[match[4]:match[5]])

		if isClosing {
			switch tagName {
			case "h1", "h2", "h3", "h4", "h5", "h6":
				fontStyle = ""
				fontSize = 12
				pdf.SetFont("Arial", fontStyle, fontSize)
				pdf.Ln(8)
			case "strong", "b":
				fontStyle = ""
				if !inCell {
					pdf.SetFont("Arial", fontStyle, fontSize)
				}
			case "em", "i":
				fontStyle = ""
				if !inCell {
					pdf.SetFont("Arial", fontStyle, fontSize)
				}
			case "code":
				fontSize = 12
				if !inCell {
					pdf.SetFont("Arial", fontStyle, fontSize)
				}
			case "p":
				if listDepth == 0 && !inTable {
					pdf.Ln(6)
				} else if !inTable {
					pdf.Ln(2)
				}
			case "li":
				pdf.SetLeftMargin(baseLeftMargin + float64(listDepth)*10)
				pdf.Ln(2)
			case "ul", "ol":
				listDepth--
				if listDepth < 0 {
					listDepth = 0
				}
				newMargin := baseLeftMargin + float64(listDepth)*10
				pdf.SetLeftMargin(newMargin)
				pdf.SetX(newMargin)
				pdf.Ln(2)
			case "blockquote":
				pdf.SetLeftMargin(baseLeftMargin)
				pdf.Ln(4)
			case "table":
				inTable = false
				pdf.SetLeftMargin(baseLeftMargin)
				pdf.SetX(baseLeftMargin)
				pdf.Ln(4)
			case "thead":
				inTableHeader = false
			case "tbody":
			case "tr":
				if len(tableRow) > 0 {
					renderTableRow(pdf, tableRow, tableColWidths, inTableHeader, baseLeftMargin)
				}
				tableRow = []tableCell{}
			case "th", "td":
				tableRow = append(tableRow, tableCell{text: strings.TrimSpace(cellText), style: cellStyle})
				cellText = ""
				cellStyle = ""
				inCell = false
			}
			justStartedListItem = false
		} else {
			switch tagName {
			case "h1":
				fontSize = 24
				fontStyle = "B"
				pdf.Ln(8)
				pdf.SetFont("Arial", fontStyle, fontSize)
			case "h2":
				fontSize = 20
				fontStyle = "B"
				pdf.Ln(6)
				pdf.SetFont("Arial", fontStyle, fontSize)
			case "h3":
				fontSize = 16
				fontStyle = "B"
				pdf.Ln(5)
				pdf.SetFont("Arial", fontStyle, fontSize)
			case "h4", "h5", "h6":
				fontSize = 14
				fontStyle = "B"
				pdf.Ln(4)
				pdf.SetFont("Arial", fontStyle, fontSize)
			case "strong", "b":
				fontStyle = "B"
				if inCell {
					cellStyle = "B"
				} else {
					pdf.SetFont("Arial", fontStyle, fontSize)
				}
			case "em", "i":
				fontStyle = "I"
				if inCell {
					cellStyle = "I"
				} else {
					pdf.SetFont("Arial", fontStyle, fontSize)
				}
			case "code":
				fontSize = 10
				if !inCell {
					pdf.SetFont("Courier", "", fontSize)
				}
			case "p":
				if !justStartedListItem && !inTable {
					pdf.Ln(4)
				}
			case "br":
				if !inTable {
					pdf.Ln(5)
				}
			case "li":
				pdf.Ln(5)
				currentMargin := baseLeftMargin + float64(listDepth)*10
				pdf.SetX(currentMargin)
				x := pdf.GetX()
				y := pdf.GetY() + 2
				pdf.SetFillColor(0, 0, 0)
				if listDepth > 1 {
					pdf.Circle(x, y, 0.8, "D")
				} else {
					pdf.Circle(x, y, 1.0, "F")
				}
				textIndent := currentMargin + 4
				pdf.SetLeftMargin(textIndent)
				pdf.SetX(textIndent)
				justStartedListItem = true
			case "ul", "ol":
				listDepth++
				pdf.SetLeftMargin(baseLeftMargin + float64(listDepth)*10)
				pdf.Ln(2)
			case "blockquote":
				pdf.SetLeftMargin(baseLeftMargin + 10)
				pdf.Ln(4)
			case "hr":
				pdf.Ln(5)
				pdf.Line(15, pdf.GetY(), 195, pdf.GetY())
				pdf.Ln(5)
			case "table":
				inTable = true
				pdf.Ln(4)
				tableColWidths = []float64{120, 60}
			case "thead":
				inTableHeader = true
			case "tbody":
				inTableHeader = false
			case "tr":
				tableRow = []tableCell{}
			case "th", "td":
				cellText = ""
				cellStyle = ""
				inCell = true
			}
		}
	}

	if lastEnd < len(htmlContent) {
		text := strings.TrimSpace(htmlContent[lastEnd:])
		if text != "" {
			text = decodeHTMLEntities(text)
			pdf.Write(5, text)
		}
	}
}

type tableCell struct {
	text  string
	style string
}

func renderTableRow(pdf *gofpdf.Fpdf, row []tableCell, colWidths []float64, isHeader bool, leftMargin float64) {
	pdf.SetX(leftMargin)

	for len(colWidths) < len(row) {
		colWidths = append(colWidths, 40)
	}

	pageWidth := 210.0 - leftMargin - 15
	totalWidth := 0.0
	for i := 0; i < len(row); i++ {
		totalWidth += colWidths[i]
	}
	if totalWidth > pageWidth {
		scale := pageWidth / totalWidth
		for i := range colWidths {
			colWidths[i] *= scale
		}
	}

	cellHeight := 8.0
	for i, cell := range row {
		width := colWidths[i]
		style := cell.style
		if isHeader && style == "" {
			style = "B"
		}
		pdf.SetFont("Arial", style, 11)
		if isHeader {
			pdf.SetFillColor(230, 230, 230)
		} else {
			pdf.SetFillColor(255, 255, 255)
		}
		pdf.CellFormat(width, cellHeight, cell.text, "1", 0, "L", isHeader, 0, "")
	}
	pdf.Ln(cellHeight)
}

func decodeHTMLEntities(text string) string {
	text = strings.ReplaceAll(text, "&nbsp;", " ")
	text = strings.ReplaceAll(text, "&amp;", "&")
	text = strings.ReplaceAll(text, "&lt;", "<")
	text = strings.ReplaceAll(text, "&gt;", ">")
	text = strings.ReplaceAll(text, "&quot;", `"`)
	text = strings.ReplaceAll(text, "&#39;", "'")
	text = strings.ReplaceAll(text, "&apos;", "'")
	text = strings.ReplaceAll(text, "&rsquo;", "'")
	text = strings.ReplaceAll(text, "&lsquo;", "'")
	text = strings.ReplaceAll(text, "&rdquo;", `"`)
	text = strings.ReplaceAll(text, "&ldquo;", `"`)
	text = strings.ReplaceAll(text, "&ndash;", "-")
	text = strings.ReplaceAll(text, "&mdash;", "-")

	numericRe := regexp.MustCompile(`&#(\d+);`)
	text = numericRe.ReplaceAllStringFunc(text, func(match string) string {
		numStr := match[2 : len(match)-1]
		var num int
		fmt.Sscanf(numStr, "%d", &num)
		if num > 0 && num < 128 {
			return string(rune(num))
		}
		return match
	})

	return text
}

func init() {
	rootCmd.AddCommand(mdCmd)
	mdCmd.AddCommand(mdToPdfCmd)
}
