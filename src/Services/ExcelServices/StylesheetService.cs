﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Services.Interfaces.ExcelServices;
using System;
using System.Linq;

namespace JsonToWord.Services.ExcelServices
{
    public class StylesheetService : IStylesheetService
    {
        public Stylesheet CreateStylesheet()
        {
            return new Stylesheet(
                new Fonts(
                new Font( // Index 0 - Default font
                        new FontSize { Val = 10 },
                        new FontName { Val = "Arial" }
                    ),
                    new Font( // Index 1 - Header font
                        new FontSize { Val = 11 },
                        new FontName { Val = "Arial" },
                        new Bold(),
                        new Color { Rgb = new HexBinaryValue("FFFFFFFF") }
                    ),
                    new Font( // Index 2 - SuiteName title font
                        new FontSize { Val = 11 },
                        new FontName { Val = "Arial" },
                        new Bold()
                    ),
                     new Font( // Index 3 - Hyperlink font
                        new FontSize { Val = 10 },
                        new FontName { Val = "Arial" },
                        new Underline { Val = UnderlineValues.Single },
                        new Color { Rgb = new HexBinaryValue("FF0563C1") } // Excel blue hyperlink color
                    )
                ),
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }), // Index 0 - Default fill
                    new Fill(new PatternFill { PatternType = PatternValues.None }), // Index 1 - Not working
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FF000000") }) { PatternType = PatternValues.Solid }), // Index 2 - Black fill for headers
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FF0E2841") }) { PatternType = PatternValues.Solid }), // Index 3 - SuiteName title fill
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FFA6C9EC") }) { PatternType = PatternValues.Solid }), // Index 4 - First alternating color
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FFDAE9F8") }) { PatternType = PatternValues.Solid }), // Index 5 - Second alternating color
                    // New Fills for Group Headers
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FF004B50") }) { PatternType = PatternValues.Solid }), // Index 6 - Test Cases Group Header
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FF0098C7") }) { PatternType = PatternValues.Solid }), // Index 7 - Requirements Group Header
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FFCC293D") }) { PatternType = PatternValues.Solid }), // Index 8 - Bugs Group Header
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FFB4009E") }) { PatternType = PatternValues.Solid })  // Index 9 - CRs Group Header
                ),
                new Borders(
                    new Border(), // Index 0 - Default border
                    new Border( // Index 1 - Thin border
                        new LeftBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin }
                    )
                ),
                new CellFormats(
                    new CellFormat(), // Index 0 - Default cell format
                    new CellFormat // Index 1 - Header format
                    {
                        FontId = 1,
                        FillId = 2,
                        BorderId = 1,
                        ApplyFont = true,
                        ApplyFill = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    },
                    new CellFormat // Index 2 - SuiteName title format
                    {
                        FontId = 1,
                        FillId = 3,
                        BorderId = 1,
                        ApplyFont = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { Vertical = VerticalAlignmentValues.Center }
                    },
                    new CellFormat(), // Index 3 - Reserved
                    new CellFormat(), // Index 4 - Reserved
                    new CellFormat(), // Index 5 - Reserved
                    new CellFormat // Index 6 - Data cell with first alternating color
                    {
                        BorderId = 1,
                        FillId = 4,
                        ApplyFill = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { WrapText = true, Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 7 - Data cell with second alternating color
                    {
                        BorderId = 1,
                        FillId = 5,
                        ApplyFill = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { WrapText = true, Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 8 - Date cell with first alternating color
                    {
                        BorderId = 1,
                        FillId = 4,
                        ApplyFill = true,
                        ApplyBorder = true,
                        NumberFormatId = 14, // Standard date format
                        ApplyNumberFormat = true,
                        Alignment = new Alignment { Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 9 - Date cell with second alternating color
                    {
                        BorderId = 1,
                        FillId = 5,
                        ApplyFill = true,
                        ApplyBorder = true,
                        NumberFormatId = 14, // Standard date format
                        ApplyNumberFormat = true,
                        Alignment = new Alignment { Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 10 - Number cell with first alternating color
                    {
                        BorderId = 1,
                        FillId = 4,
                        ApplyFill = true,
                        ApplyBorder = true,
                        NumberFormatId = 0, // General number format
                        ApplyNumberFormat = true,
                        Alignment = new Alignment { Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat { FontId = 0, FillId = 5, BorderId = 1, NumberFormatId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyNumberFormat = true }, // Index 11 - Number with alternating color 2

                    new CellFormat // Index 12 - Hyperlink style for first alternating color
                    {
                        FontId = 3, // Use the hyperlink font we defined
                        FillId = 4, // First alternating color fill
                        BorderId = 1,
                        ApplyFont = true,
                        ApplyFill = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { WrapText = true, Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 13 - Hyperlink style for second alternating color
                    {
                        FontId = 3, // Use the hyperlink font we defined
                        FillId = 5, // Second alternating color fill
                        BorderId = 1,
                        ApplyFont = true,
                        ApplyFill = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { WrapText = true, Vertical = VerticalAlignmentValues.Top }
                    },

                    // New CellFormats for Group Headers - Moved to a higher index to avoid conflicts
                    new CellFormat { FontId = 1, FillId = 6, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }, ApplyAlignment = true }, // Index 14 - Test Cases Group Header
                    new CellFormat { FontId = 1, FillId = 7, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }, ApplyAlignment = true }, // Index 15 - Requirements Group Header
                    new CellFormat { FontId = 1, FillId = 8, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }, ApplyAlignment = true }, // Index 16 - Bugs Group Header
                    new CellFormat { FontId = 1, FillId = 9, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }, ApplyAlignment = true }  // Index 17 - CRs Group Header
                )
            );
        }

        public void EnsureStylesheet(WorkbookPart workbookPart)
        {
            if (workbookPart.GetPartsOfType<WorkbookStylesPart>().Any())
                return;

            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateStylesheet();
            stylesPart.Stylesheet.Save();
        }

    }
}
