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

var mdToPdfCmd = &cobra.Command{
	Use:   "mdToPdf [file]",
	Short: "Convert markdown file to PDF",
	Long:  "Convert a markdown file to a PDF file. Takes a markdown file path as argument.",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		inputPath := args[0]

		// Check if input file exists
		if _, err := os.Stat(inputPath); os.IsNotExist(err) {
			fmt.Printf("Error: File not found: %s\n", inputPath)
			return
		}

		// Read markdown file
		mdContent, err := os.ReadFile(inputPath)
		if err != nil {
			fmt.Printf("Error reading file: %v\n", err)
			return
		}

		// Create output filename
		ext := filepath.Ext(inputPath)
		baseName := strings.TrimSuffix(inputPath, ext)
		outputPath := baseName + ".pdf"

		// Convert markdown to HTML
		extensions := parser.CommonExtensions | parser.AutoHeadingIDs
		p := parser.NewWithExtensions(extensions)
		htmlContent := markdown.ToHTML(mdContent, p, nil)

		// Create PDF
		pdf := gofpdf.New("P", "mm", "A4", "")
		pdf.SetMargins(15, 15, 15)
		pdf.AddPage()
		pdf.SetFont("Arial", "", 12)

		// Process HTML content
		renderHTMLToPDF(pdf, string(htmlContent))

		// Save PDF
		if err := pdf.OutputFileAndClose(outputPath); err != nil {
			fmt.Printf("Error creating PDF: %v\n", err)
			return
		}

		fmt.Printf("PDF created successfully: %s\n", outputPath)
	},
}

func renderHTMLToPDF(pdf *gofpdf.Fpdf, htmlContent string) {
	// Remove comments
	htmlContent = regexp.MustCompile(`<!--.*?-->`).ReplaceAllString(htmlContent, "")

	// Parse HTML tags and text content using regex
	tagRe := regexp.MustCompile(`<(/?)([a-zA-Z0-9]+)[^>]*>`)

	fontStyle := ""
	fontSize := 12.0
	listDepth := 0
	baseLeftMargin := 15.0
	justStartedListItem := false

	// Table state
	inTable := false
	inTableHeader := false
	tableRow := []tableCell{}
	tableColWidths := []float64{}
	cellText := ""
	cellStyle := ""
	inCell := false

	// Find all tag positions
	lastEnd := 0
	matches := tagRe.FindAllStringSubmatchIndex(htmlContent, -1)

	for _, match := range matches {
		// Get text before this tag
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

		// Extract tag info
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
				// Don't reset cellStyle here - keep it for the cell content
				if !inCell {
					pdf.SetFont("Arial", fontStyle, fontSize)
				}
			case "em", "i":
				fontStyle = ""
				// Don't reset cellStyle here - keep it for the cell content
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
				// Reset margin to list level (not text indent)
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
				// Nothing special
			case "tr":
				// Render the row
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
				// Skip line break if we just started a list item
				if !justStartedListItem && !inTable {
					pdf.Ln(4)
				}
			case "br":
				if !inTable {
					pdf.Ln(5)
				}
			case "li":
				pdf.Ln(5)
				// Set X to current left margin for proper bullet placement
				currentMargin := baseLeftMargin + float64(listDepth)*10
				pdf.SetX(currentMargin)
				// Draw a bullet circle
				x := pdf.GetX()
				y := pdf.GetY() + 2
				pdf.SetFillColor(0, 0, 0) // Black fill for bullet
				if listDepth > 1 {
					// Nested: outline circle (empty inside)
					pdf.Circle(x, y, 0.8, "D")
				} else {
					// Top level: filled circle
					pdf.Circle(x, y, 1.0, "F")
				}
				// Set margin for text wrap alignment (after bullet)
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
				// Default column widths - will be adjusted based on content
				tableColWidths = []float64{120, 60} // Adjust as needed
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

	// Handle any remaining text after the last tag
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

	// Adjust column widths if we have more columns than defined
	for len(colWidths) < len(row) {
		colWidths = append(colWidths, 40)
	}

	// Calculate total width and adjust if needed
	pageWidth := 210.0 - leftMargin - 15 // A4 width minus margins
	totalWidth := 0.0
	for i := 0; i < len(row); i++ {
		totalWidth += colWidths[i]
	}

	// Scale column widths to fit page
	if totalWidth > pageWidth {
		scale := pageWidth / totalWidth
		for i := range colWidths {
			colWidths[i] *= scale
		}
	}

	cellHeight := 8.0

	for i, cell := range row {
		width := colWidths[i]

		// Determine font style
		style := cell.style
		if isHeader && style == "" {
			style = "B"
		}
		pdf.SetFont("Arial", style, 11)

		// Set fill color
		if isHeader {
			pdf.SetFillColor(230, 230, 230)
		} else {
			pdf.SetFillColor(255, 255, 255)
		}

		// Draw cell with border
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

	// Handle numeric character references
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
	rootCmd.AddCommand(mdToPdfCmd)
}
