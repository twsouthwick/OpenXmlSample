// Licensed under the MIT license. See LICENSE file in the samples root for full license information.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using System.Reflection;

using static OpenXml.OpenXmlConstants;

namespace OpenXml
{
    internal static class OpenXmlExtensions
    {
        /// <summary>
        /// Adds a new sheet into the spreadsheet's workbook
        /// </summary>
        /// <param name="spreadsheet">The spreadsheetdocument object</param>
        /// <param name="name">Name to call the spreadsheet</param>
        /// <returns>The newly added spreadsheet</returns>
        public static Worksheet AddWorksheet(this SpreadsheetDocument spreadsheet, string name)
        {
            // Get the spreadsheets collection
            var sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheets == null)
            {
                sheets = new Sheets();
                spreadsheet.WorkbookPart.Workbook.AppendChild(sheets);
            }

            // Add new worksheet
            var worksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();
            worksheetPart.Worksheet.Save();

            // Create the worksheet to workbook relation
            sheets.AppendChild(new Sheet()
            {
                Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = new UInt32Value((uint)sheets.Count() + 1),
                Name = name
            });

            return worksheetPart.Worksheet;
        }

        /// <summary>
        /// Adds the predefined xml style content from embedded .xml file to the workbook
        /// </summary>
        /// <param name="workBookPart">Workbookpart to add styles to</param>
        public static void AddPreDefinedSpreadsheetStyles(this WorkbookPart workBookPart)
        {
            var styleSheet = new Stylesheet()
            {
                // This method pulls the PreDefinedStyles.xml as an embeded file
                InnerXml = GetXMLFromEmbededResource("PreDefinedStyles.xml")
            };

            ValidateStylesheet(styleSheet);

            workBookPart.AddNewPart<WorkbookStylesPart>();
            workBookPart.WorkbookStylesPart.Stylesheet = styleSheet;
        }

        /// <summary>
        /// Uses the openxml object model to add styles to a spreadsheet in the workbook
        /// </summary>
        /// <param name="workBookPart">The workbook part to add the style to</param>
        public static void AddSpreadsheetStylesThroughCode(this WorkbookPart workBookPart)
        {
            var styleSheet = new Stylesheet()
            {
                MCAttributes = new MarkupCompatibilityAttributes { Ignorable = $"{X14ac} {X16r2}" },
                Borders = BuildBorders(),
                CellFormats = BuildCellFormats(),
                CellStyleFormats = BuildCellStyleFormats(),
                CellStyles = BuildCellStyles(),
                Fills = BuildFills(),
                Fonts = BuildFonts(),
            };

            AddSchemaNamespaces(styleSheet);

            ValidateStylesheet(styleSheet);

            workBookPart.AddNewPart<WorkbookStylesPart>();
            workBookPart.WorkbookStylesPart.Stylesheet = styleSheet;
        }

        /// <summary>
        /// Adds a row to the worksheet
        /// </summary>
        /// <param name="workSheet">Worksheet to add to</param>
        /// <param name="data">An array of objects to add</param>
        /// <param name="styleIndex">Style index from the workbooks styles to use</param>
        public static void AddRow(this Worksheet workSheet, object[] data, uint styleIndex = 0)
        {
            var sd = workSheet.GetFirstChild<SheetData>();
            if (sd == null)
            {
                sd = workSheet.AppendChild(new SheetData());
            }

            var row = sd.AppendChild(new Row());

            foreach (var item in data)
            {
                Cell GetCell()
                {
                    switch (item)
                    {
                        case null:
                            return new Cell();
                        case DateTimeOffset dt:
                            return BuildDateCell(dt);
                        case int i:
                            return BuildNumberCell(i.ToString());
                        case PropertyInfo pi:
                            return BuildTextCell(pi.Name);
                        case string str:
                            return BuildTextCell(str);
                        default:
                            Console.WriteLine("Warning: item is not a string. The item will be written with .ToString()");
                            return BuildTextCell(item.ToString());
                    }
                }

                var cell = GetCell();
                cell.StyleIndex = styleIndex;
                row.AppendChild(cell);
            }
        }

        /// <summary>
        /// Builds a string of an embedded excel style sheet
        /// </summary>
        /// <param name="name">Name of embeded resource to use</param>
        /// <returns>Strng of Excel styling information</returns>
        private static string GetXMLFromEmbededResource(string name)
        {
            var assembly = typeof(OpenXmlExtensions).GetTypeInfo().Assembly;
            var names = assembly.GetManifestResourceNames();
            var resourceName = $"{assembly.GetName().Name}.Resources.{name}";

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new ArgumentException($"Embedded resource {resourceName} could not be found!");
                }

                using (var streamReader = new StreamReader(stream))
                {
                    return streamReader.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Builds an excel text cell
        /// </summary>
        /// <param name="text">The value to put in the cell</param>
        /// <returns>Excel text cell</returns>
        private static Cell BuildTextCell(string text)
        {
            var cell = new Cell
            {
                DataType = CellValues.InlineString,
            };

            var inlineString = new InlineString
            {
                Text = new Text(text)
            };

            cell.AppendChild(inlineString);

            return cell;
        }

        /// <summary>
        /// Builds an excel number cell
        /// </summary>
        /// <param name="value">Value to add to the cell</param>
        /// <returns>Excel number cell</returns>
        private static Cell BuildNumberCell(string value)
        {
            return new Cell
            {
                CellValue = new CellValue(value)
            };
        }

        /// <summary>
        /// Builds an excel column showing booling text values (True/False)
        /// </summary>
        /// <param name="value">The boolean value to use</param>
        /// <returns>Excel boolean column</returns>
        private static Cell BuildBooleanCell(bool value)
        {
            return BuildTextCell(value.ToString());
        }

        /// <summary>
        /// Builds an excel text cell containing a date
        /// </summary>
        /// <param name="value">The date value to use</param>
        /// <returns>Excel text cell</returns>
        private static Cell BuildDateCell(DateTimeOffset value)
        {
            return BuildTextCell(value.Date.ToString("MMM dd, yyyy"));
        }

        /// <summary>
        /// Validates an excel style sheet
        /// </summary>
        /// <param name="styleSheet">The style sheet to validate</param>
        private static void ValidateStylesheet(Stylesheet styleSheet)
        {
            var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
            var validationErrors = validator.Validate(styleSheet).ToList();

            if (validationErrors.Count > 0)
            {
                Console.WriteLine($"There were validation errors with the style sheet {string.Join(Environment.NewLine, validationErrors.Select(r => r.Description))}");
            }
        }

        /// <summary>
        /// Adds schema values into the style sheet
        /// </summary>
        /// <param name="styleSheet">Style sheet to add to</param>
        private static void AddSchemaNamespaces(Stylesheet styleSheet)
        {
            styleSheet.AddNamespaceDeclaration(Mc, @"http://schemas.openxmlformats.org/markup-compatibility/2006");
            styleSheet.AddNamespaceDeclaration(X14ac, @"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            styleSheet.AddNamespaceDeclaration(X16r2, @"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
        }

        /// <summary>
        /// Builds a collection of cellstyles
        /// </summary>
        /// <returns>Collection of cellstyles</returns>
        private static CellStyles BuildCellStyles()
        {
            return new CellStyles(
                new CellStyle
                {
                    Name = "Normal",
                    FormatId = 0U,
                    BuiltinId = 0U
                });
        }

        /// <summary>
        /// Builds a collection of cellformats
        /// </summary>
        /// <returns>Collection of cellformats</returns>
        private static CellFormats BuildCellFormats()
        {
            return new CellFormats(
                BuildCellFormat(numberFormatId: 0, fontId: 0, fillId: 0, borderId: 0, formatId: 0, alignment: new Alignment()),
                BuildCellFormat(numberFormatId: 0, fontId: 0, fillId: 0, borderId: 0, formatId: 0, applyAlignment: true,
                                alignment: new Alignment() { Horizontal = HorizontalAlignmentValues.Center }),
                BuildCellFormat(numberFormatId: 0, fontId: 1, fillId: 2, borderId: 0, formatId: 0,
                                alignment: new Alignment(), applyFont: true, applyFill: true),
                BuildCellFormat(numberFormatId: 0, fontId: 1, fillId: 2, borderId: 0, formatId: 0, applyFont: true, applyFill: true,
                                applyAlignment: true, alignment: new Alignment() { Horizontal = HorizontalAlignmentValues.Center }));
        }

        /// <summary>
        /// Builds a cellformat object using the passed in values
        /// </summary>
        /// <returns>Cellformat object</returns>
        private static CellFormat BuildCellFormat(UInt32Value numberFormatId,
                                                  UInt32Value fontId,
                                                  UInt32Value fillId,
                                                  UInt32Value borderId,
                                                  UInt32Value formatId,
                                                  Alignment alignment,
                                                  bool applyFont = false,
                                                  bool applyFill = false,
                                                  bool applyAlignment = false)
        {
            return new CellFormat
            {
                NumberFormatId = numberFormatId,
                FontId = fontId,
                FillId = fillId,
                BorderId = borderId,
                FormatId = formatId,
                ApplyFill = applyFill,
                ApplyFont = applyFont,
                ApplyAlignment = applyAlignment,
                Alignment = alignment
            };
        }

        /// <summary>
        /// Builds a cellstyleformats object
        /// </summary>
        /// <returns>Cellstylesformats object</returns>
        private static CellStyleFormats BuildCellStyleFormats()
        {
            return new CellStyleFormats(
                new CellFormat()
                {
                    BorderId = 0U,
                    FillId = 0U,
                    FontId = 0U,
                    NumberFormatId = 0U
                });
        }

        /// <summary>
        /// Builds a borders object
        /// </summary>
        /// <returns>Borders object</returns>
        private static Borders BuildBorders()
        {
            return new Borders(
                new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()));
        }

        /// <summary>
        /// Builds a fills object
        /// </summary>
        /// <returns>Fills object</returns>
        private static Fills BuildFills()
        {
            return new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                new Fill(
                    new PatternFill
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor { Theme = 4U, Tint = 0.59999389629810485D },
                        BackgroundColor = new BackgroundColor { Indexed = 64U }
                    }));
        }

        /// <summary>
        /// Builds a fonts collection
        /// </summary>
        /// <returns>Fonts collection</returns>
        private static Fonts BuildFonts()
        {
            return new Fonts(
                new Font(
                    new FontSize() { Val = 11D },
                    new Color() { Theme = 1U },
                    new FontName() { Val = "Calibri" },
                    new FontFamilyNumbering() { Val = 2 },
                    new FontScheme() { Val = FontSchemeValues.Minor }),
                new Font(
                    new Bold(),
                    new FontSize() { Val = 14D },
                    new Color() { Theme = 1U },
                    new FontName() { Val = "Calibri" },
                    new FontFamilyNumbering() { Val = 2 },
                    new FontScheme() { Val = FontSchemeValues.Minor }));
        }
    }
}
