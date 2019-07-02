using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteExcelSAX
{
    public class Report
    {
        private static readonly string FileName = @"D:\test.xlsx";

        public void WriteDataSAX()
        {
            // (1) Create a file (Document)
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(FileName, SpreadsheetDocumentType.Workbook)) {
                // (2) Add Workbook to Doc
                WorkbookPart workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                // (3) Add a style sheet
                WorkbookStylesPart wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();
                // add styles to sheet
                wbsp.Stylesheet = GenerateStylesheet();
                wbsp.Stylesheet.Save();
                // (4) Add Sheets to Workbook
                Sheets sheets = doc.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                // (5.1) Add the first Worksheet
                var workSheetPart1 = AddWorksheet(doc, 1, "My Page 1");

                WriteData(workSheetPart1);
                doc.Save();
                doc.Close();
            }
        }

        private static void WriteData(WorksheetPart workSheetPart)
        {
            OpenXmlWriter writer = OpenXmlWriter.Create(workSheetPart);
            {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                // Write Values
                for (int row = 0; row < 2000; row++) {
                    Row r = new Row() {
                        RowIndex = Convert.ToUInt32(row + 2)
                    };
                    writer.WriteStartElement(r);

                    for (int col = 0; col < 20; col++) {
                        Cell c = new Cell() {
                            StyleIndex = Convert.ToUInt32(2),
                            CellReference = new StringValue()
                        };
                        CellValue v = new CellValue("xxx");

                        c.DataType = new EnumValue<CellValues>(CellValues.String);
                        c.Append(v);
                        writer.WriteElement(c);
                    }
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndElement();

                writer.Close();
            }
        }

        private WorksheetPart AddWorksheet(SpreadsheetDocument doc, int sheetId, string sheetName)
        {
            // (1) Add WorksheetPart to WorkbookPart
            WorksheetPart worksheetPart = doc.WorkbookPart.AddNewPart<WorksheetPart>();

            // (2) Get a worksheetId
            var workSheetPartId = doc.WorkbookPart.GetIdOfPart(worksheetPart);

            // (3) Create a shee
            Sheet sheet = new Sheet() {
                Id = workSheetPartId,
                SheetId = Convert.ToUInt32(sheetId),
                Name = sheetName
            };
            doc.WorkbookPart.Workbook.Sheets.Append(sheet);

            return worksheetPart;
        }

        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }

                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "990000" } }) { PatternType = PatternValues.Solid }), // Index 2 - header
                     new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "666666" } }) { PatternType = PatternValues.Solid })

                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true, Rgb = "FF0000" }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new Border(),
                        new DiagonalBorder())
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true },// header
                      new CellFormat { FontId = 1, FillId = 3, BorderId = 1, ApplyFill = true } // header
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }
    }
}