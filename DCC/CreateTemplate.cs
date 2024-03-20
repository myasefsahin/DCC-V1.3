using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;
using System.Windows.Forms;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableStyle = DocumentFormat.OpenXml.Wordprocessing.TableStyle;
using DCC;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace DCC
{
    class CreateTemplate
    {
       
        #region CreateForTwoRow
        public Table CreateForTwoRow(ArrayList Freq, ArrayList Value1, ArrayList Value1Unc, String Title, String TitleEng)
        {
            // Table properties and some other specs
            Table table = new Table(
                new TableProperties(
                    new TableStyle() { Val = "TableGrid" },
                    new TableWidth() { Type = TableWidthUnitValues.Auto },
                    new TableLook() { Val = "04A0" }
                ),
                new TableGrid(
                    new GridColumn() { Width = "833" } // Define the width for each column
                ),
                new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = 9 },
                    new BottomBorder() { Val = BorderValues.Single, Size = 9 },
                    new LeftBorder() { Val = BorderValues.Single, Size = 9 },
                    new RightBorder() { Val = BorderValues.Single, Size = 9 },
                    new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 9 },
                    new InsideVerticalBorder() { Val = BorderValues.Single, Size = 9 }
                ),
                new TableRow(

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Frekans (GHz)"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Frequency (GHz)")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "4430" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text(Title))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(new RunProperties(new Italic()), new Text(TitleEng)))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "4430" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text(Title + " Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text(TitleEng + " Uncertainty"))))
                ),
                new TableJustification() { Val = TableRowAlignmentValues.Center }
            // Continue adding rows here...
            );

            // Set cell margins, align text vertically, and align text horizontally
            foreach (TableRow row in table.Elements<TableRow>())
            {
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    TableCellProperties cellProperties = cell.GetFirstChild<TableCellProperties>();
                    if (cellProperties == null)
                    {
                        cellProperties = new TableCellProperties();
                        cell.AppendChild(cellProperties);
                    }

                    // Set cell margins
                    TableCellMargin cellMargin = new TableCellMargin(
                        new LeftMargin() { Width = "69", Type = TableWidthUnitValues.Dxa },
                        new RightMargin() { Width = "69", Type = TableWidthUnitValues.Dxa },
                        new TopMargin() { Width = "50", Type = TableWidthUnitValues.Dxa },
                        new BottomMargin() { Width = "50", Type = TableWidthUnitValues.Dxa }
                    );
                    cellProperties.Append(cellMargin);

                    // Align text vertically
                    cellProperties.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });

                    // Align text horizontally
                    foreach (Paragraph paragraph in cell.Elements<Paragraph>())
                    {
                        ParagraphProperties paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>();
                        if (paragraphProperties == null)
                        {
                            paragraphProperties = new ParagraphProperties();
                            paragraph.InsertAt(paragraphProperties, 0);
                        }

                        // Set text alignment
                        paragraphProperties.Append(new Justification() { Val = JustificationValues.Center }); // Horizontally center
                    }
                }
            }

            // Apply bold formatting to the runs in the first row
            //TableRow firstRow = table.Elements<TableRow>().FirstOrDefault();
            //if (firstRow != null) {
            //    foreach (Run run in firstRow.Descendants<Run>()) {
            //        run.RunProperties = new RunProperties(new Bold());
            //    }
            //}

            // Import data from Form Interface for Reel & Imaginary
            for (int i = 0; i < Freq.Count; i++)
            {

                //double swrData = (double)System.Math.Round(Convert.ToDouble(Swr[i]), 4);
                //double swruncData = (double)System.Math.Round(Convert.ToDouble(SwrUnc[i]), 4);

                TableRow row = new TableRow(

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Freq[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Value1[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Value1Unc[i].ToString()))))
                );

                // Set cell margins, align text vertically, and align text horizontally
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    TableCellProperties cellProperties = cell.GetFirstChild<TableCellProperties>();
                    if (cellProperties == null)
                    {
                        cellProperties = new TableCellProperties();
                        cell.AppendChild(cellProperties);
                    }

                    // Set cell margins
                    TableCellMargin cellMargin = new TableCellMargin(
                        new LeftMargin() { Width = "69", Type = TableWidthUnitValues.Dxa },
                        new RightMargin() { Width = "69", Type = TableWidthUnitValues.Dxa },
                        new TopMargin() { Width = "50", Type = TableWidthUnitValues.Dxa },
                        new BottomMargin() { Width = "50", Type = TableWidthUnitValues.Dxa }
                    );
                    cellProperties.Append(cellMargin);

                    // Align text vertically
                    cellProperties.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });

                    // Align text horizontally
                    foreach (Paragraph paragraph in cell.Elements<Paragraph>())
                    {
                        ParagraphProperties paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>();
                        if (paragraphProperties == null)
                        {
                            paragraphProperties = new ParagraphProperties();
                            paragraph.InsertAt(paragraphProperties, 0);
                        }

                        // Set text alignment
                        paragraphProperties.Append(new Justification() { Val = JustificationValues.Center }); // Horizontally center
                    }
                }
                table.Append(row);
            }
            return table;
        }
        #endregion
        #region CreateForFourRow
        public Table CreateForFourRow(ArrayList Freq, ArrayList Value1, ArrayList Value1Unc, ArrayList Value2, ArrayList Value2Unc, String Title1, String Title1Eng, String Title2, String Title2Eng)
        {
            // Table properties and some other specs
            Table table = new Table(
                new TableProperties(
                    new TableStyle() { Val = "TableGrid" },
                    new TableWidth() { Type = TableWidthUnitValues.Auto },
                    new TableLook() { Val = "04A0" }
                ),
                new TableGrid(
                    new GridColumn() { Width = "500" } // Define the width for each column
                ),
                new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = 9 },
                    new BottomBorder() { Val = BorderValues.Single, Size = 9 },
                    new LeftBorder() { Val = BorderValues.Single, Size = 9 },
                    new RightBorder() { Val = BorderValues.Single, Size = 9 },
                    new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 9 },
                    new InsideVerticalBorder() { Val = BorderValues.Single, Size = 9 }
                ),
                new TableRow(

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Frekans (GHz)"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Frequency (GHz)")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text(Title1))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text(Title1Eng)))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text(Title2 + " Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text(Title2Eng + " Uncertainty")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text(Title2))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text(Title2)))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text(Title2 + " Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text(Title2 + " Uncertainty"))))
                ),
                new TableJustification() { Val = TableRowAlignmentValues.Center }
            // Continue adding rows here...
            );

            // Set cell margins, align text vertically, and align text horizontally
            foreach (TableRow row in table.Elements<TableRow>())
            {
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    TableCellProperties cellProperties = cell.GetFirstChild<TableCellProperties>();
                    if (cellProperties == null)
                    {
                        cellProperties = new TableCellProperties();
                        cell.AppendChild(cellProperties);
                    }

                    // Set cell margins
                    TableCellMargin cellMargin = new TableCellMargin(
                        new LeftMargin() { Width = "69", Type = TableWidthUnitValues.Dxa },
                        new RightMargin() { Width = "69", Type = TableWidthUnitValues.Dxa },
                        new TopMargin() { Width = "50", Type = TableWidthUnitValues.Dxa },
                        new BottomMargin() { Width = "50", Type = TableWidthUnitValues.Dxa }
                    );
                    cellProperties.Append(cellMargin);

                    // Align text vertically
                    cellProperties.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });

                    // Align text horizontally
                    foreach (Paragraph paragraph in cell.Elements<Paragraph>())
                    {
                        ParagraphProperties paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>();
                        if (paragraphProperties == null)
                        {
                            paragraphProperties = new ParagraphProperties();
                            paragraph.InsertAt(paragraphProperties, 0);
                        }

                        // Set text alignment
                        paragraphProperties.Append(new Justification() { Val = JustificationValues.Center }); // Horizontally center
                    }
                }
            }

            // Apply bold formatting to the runs in the first row
            //TableRow firstRow = table.Elements<TableRow>().FirstOrDefault();
            //if (firstRow != null) {
            //    foreach (Run run in firstRow.Descendants<Run>()) {
            //        run.RunProperties = new RunProperties(new Bold());
            //    }
            //}

            // Import data from Form Interface for Reel & Imaginary
            for (int i = 0; i < Freq.Count; i++)
            {

                //double logData = (double)System.Math.Round(Convert.ToDouble(LogMag[i]), 4);
                //double loguncData = (double)System.Math.Round(Convert.ToDouble(LogMagUnc[i]), 4);
                //double fazData = (double)System.Math.Round(Convert.ToDouble(LogPhase[i]), 4);
                //double fazuncData = (double)System.Math.Round(Convert.ToDouble(LogPhaseUnc[i]), 4);

                TableRow row = new TableRow(

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Freq[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Value1[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Value1Unc[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Value2[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Value2Unc[i].ToString()))))
                );

                // Set cell margins, align text vertically, and align text horizontally
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    TableCellProperties cellProperties = cell.GetFirstChild<TableCellProperties>();
                    if (cellProperties == null)
                    {
                        cellProperties = new TableCellProperties();
                        cell.AppendChild(cellProperties);
                    }

                    // Set cell margins
                    TableCellMargin cellMargin = new TableCellMargin(
                        new LeftMargin() { Width = "69", Type = TableWidthUnitValues.Dxa },
                        new RightMargin() { Width = "69", Type = TableWidthUnitValues.Dxa },
                        new TopMargin() { Width = "50", Type = TableWidthUnitValues.Dxa },
                        new BottomMargin() { Width = "50", Type = TableWidthUnitValues.Dxa }
                    );
                    cellProperties.Append(cellMargin);

                    // Align text vertically
                    cellProperties.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });

                    // Align text horizontally
                    foreach (Paragraph paragraph in cell.Elements<Paragraph>())
                    {
                        ParagraphProperties paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>();
                        if (paragraphProperties == null)
                        {
                            paragraphProperties = new ParagraphProperties();
                            paragraph.InsertAt(paragraphProperties, 0);
                        }

                        // Set text alignment
                        paragraphProperties.Append(new Justification() { Val = JustificationValues.Center }); // Horizontally center
                    }
                }
                table.Append(row);
            }
            return table;
        }

        #endregion
        #region Cover Page
        public Table CoverPage(string Customer, string Order, string Device, string Manufacturer, string Type, string Serial, string dateOfCalibration)
        {
            Table table = new Table(
    new TableProperties(
        new TableStyle() { Val = "TableGrid" },
        new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "100%" },
        new TableLook() { Val = "04A0" },
        new TableBorders(
            new TopBorder() { Val = BorderValues.Single, Size = 6, Color = "auto", Space = 0 } // Sadece üst kenarlık
        )
    )
);


            bool isFirstRow = true;

            int i = 1;
            string FullText;
            string FullTextEng;
            // Create a paragraph for each of the Order, Device, and Serial strings
            foreach (string text in new[] { Customer, Order, Device, Manufacturer, Type, Serial, dateOfCalibration })
            {

                // If it's not the first row, add an empty paragraph for spacing
                if (!isFirstRow)
                {
                    Paragraph spacingParagraph = new Paragraph(
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Center }
                        ),
                        new Run(
                            new Text("\n") // Add a new line character for spacing
                        )
                    );
                    table.AppendChild(new TableRow(new TableCell(spacingParagraph)));
                }
                else
                {
                    isFirstRow = false;
                }

                if (i == 1)
                {
                    FullText = "Cihazın Sahibi/Adresi: " + Customer;
                    FullTextEng = "Customer/Adress: " + Customer;
                }
                else if (i == 2)
                {
                    FullText = "İstek Numarası Name: " + Order;
                    FullTextEng = "OrderNo: " + Order;
                }
                else if (i == 3)
                {
                    FullText = "Makina/Cihaz: " + Device;
                    FullTextEng = "Instrument/Device: " + Device;
                }
                else if (i == 4)
                {
                    FullText = "İmalatçı: " + Manufacturer;
                    FullTextEng = "Manufacturer: " + Manufacturer;
                }
                else if (i == 5)
                {
                    FullText = "Tip: " + Type;
                    FullTextEng = "Type: " + Type;
                }
                else if (i == 6)
                {
                    FullText = "Seri Numarası: " + Serial;
                    FullTextEng = "Serial Number: " + Serial;
                }
                else if (i == 7)
                {
                    FullText = "Kalibrasyon Tarihi: " + dateOfCalibration;
                    FullTextEng = "Date of Calibration: " + dateOfCalibration;
                }
                else
                {
                    FullText = "NULL";
                    FullTextEng = "NULL";
                }
                // Create a paragraph for the current text
                Paragraph paragraph = new Paragraph(

     new ParagraphProperties(
         new Justification() { Val = JustificationValues.Left }

     ),
     new Run(
         new RunProperties(
             new Bold(),
             new RunFonts() { Ascii = "Arial Bold", HighAnsi = "Arial Bold" },
             new FontSize() { Val = "22" }
         ),
         new Text(FullText)
     )
 );

                Paragraph englishParagraph = new Paragraph(

                    new Run(
                        new RunProperties(
                            new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                            new FontSize() { Val = "20" },
                            new Italic()
                        ),
                        new Text(FullTextEng)
                    )
                );


                // Türkçe metni içeren paragrafın sonuna alt satır karakteri ekle
                paragraph.AppendChild(new Break());

                // İngilizce metni içeren paragrafın sonuna alt satır karakteri ekle
                englishParagraph.AppendChild(new Break());

                // Paragrafları birleştir
                paragraph.Append(englishParagraph);


                // Add the formatted text to the table cell
                TableCell cell = new TableCell(paragraph);

                // Center-align the cell horizontally on the page
                TableCellProperties cellProperties = new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = "100%" }, // Set the cell width to 100% of the page width
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                );
                cell.TableCellProperties = cellProperties;

                // Add the cell to the table
                TableRow row = new TableRow(cell);
                table.Append(row);

                i++;
            }
            return table;
        }


        #endregion

        public Table UsedDevice()
        {
            Table table = new Table();

            // Define table properties
            TableProperties tableProperties = new TableProperties(
                new TableLayout { Type = TableLayoutValues.Fixed },
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "100%" }
            );
            table.Append(tableProperties);

            // Define table columns
            TableGrid tableGrid = new TableGrid(
                new GridColumn(), new GridColumn(), new GridColumn(), new GridColumn()
            );
            table.Append(tableGrid);

            TableBorders tableBorders = new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 }
            );
            tableProperties.Append(tableBorders);

            // Define table rows and cells
            TableRow row1 = new TableRow(
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("No")))
                ),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Cihaz Adı")))
                ),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Üretici Firma")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Tip/Model")))
                ),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Seri No")))
                ),
                  new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("İzlenebilirlik")))
                )
            );

            TableRowProperties rowProperties1 = new TableRowProperties(new TableJustification() { Val = TableRowAlignmentValues.Center });
            row1.Append(rowProperties1);

            table.Append(row1);

            TableRow row2 = new TableRow(
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text("Yuh"))))
                ,
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text("Yuh"))))
                ,
                        new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text("Yuh"))))
                ,
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text("Yuh")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text("Yuh")))
                ),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text("Yuh")))
                )
            );
            TableRowProperties rowProperties2 = new TableRowProperties(new TableRowHeight() { Val = 200, HeightType = HeightRuleValues.AtLeast });
            row2.Append(rowProperties2);

            table.Append(row2);

            return table;
        }
        public Table DeviceTable(string deviceName, string customer, string type, string measurement)
        {
            Table table = new Table();

            // Define table properties
            TableProperties tableProperties = new TableProperties(
                new TableLayout { Type = TableLayoutValues.Fixed },
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "100%" }
            );
            table.Append(tableProperties);

            // Define table columns
            TableGrid tableGrid = new TableGrid(
                new GridColumn(), new GridColumn(), new GridColumn(), new GridColumn()
            );
            table.Append(tableGrid);

            TableBorders tableBorders = new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 }
            );
            tableProperties.Append(tableBorders);

            // Define table rows and cells
            TableRow row1 = new TableRow(
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Cihaz Adı")))
                ),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Üretici Firma")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Tip/Model")))
                ),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Ölçüm Aralığı")))
                )
            );

            TableRowProperties rowProperties1 = new TableRowProperties(new TableJustification() { Val = TableRowAlignmentValues.Center });
            row1.Append(rowProperties1);

            table.Append(row1);

            TableRow row2 = new TableRow(
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(deviceName))))
                ,

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(customer)))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(type)))
                ),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                    new Paragraph(
                        new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                        new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(measurement)))
                )
            );
            TableRowProperties rowProperties2 = new TableRowProperties(new TableRowHeight() { Val = 200, HeightType = HeightRuleValues.AtLeast });
            row2.Append(rowProperties2);

            table.Append(row2);

            return table;
        }
        public void HeaderPage(KbysEntity kbysEntity)
        {
            string originalFilePath = "C:\\Users\\work\\Desktop\\sertifikaC.docx";
            string newFilePath = string.Empty;

            // Kullanıcıya kaydedilecek konumu seçme iletişim kutusunu göster
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = "C:\\Users\\work\\Desktop";
            saveFileDialog.Filter = "Word Dosyaları|*.docx";
            saveFileDialog.Title = "Kopyayı kaydedin";



            // Kullanıcı bir konum seçtiyse
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // İlk belgenin kopyasını seçilen konuma oluştur
                    newFilePath = saveFileDialog.FileName;
                    File.Copy(originalFilePath, newFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Kopya oluşturulurken bir hata oluştu: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Kayıt konumu seçilmedi.");
            }


            Dictionary<string, string> replacements = new Dictionary<string, string>
        {
            { "AS1", kbysEntity.CustomerName },
            { "AS2", kbysEntity.OrderNo },
            { "AS3", kbysEntity.Device },
            { "AS4", kbysEntity.Manufactuer },
            { "AS5", kbysEntity.Type },
            { "AS6", kbysEntity.SerialNumber },
            { "AS7", kbysEntity.DateOfCalibration },
            { "ASDATE", DateTime.UtcNow.ToString("dd/MM/yyyy") }
            // İhtiyacınıza göre buraya diğer değişikliklerinizi ekleyebilirsiniz
        };

            using (WordprocessingDocument doc = WordprocessingDocument.Open(newFilePath, true))
            {
                var body = doc.MainDocumentPart.Document.Body;
                foreach (var text in body.Descendants<Text>())
                {
                    foreach (var replacement in replacements)
                    {
                        text.Text = text.Text.Replace(replacement.Key, replacement.Value);
                    }
                }
            }
            MessageBox.Show("Belge başarıyla güncellendi.");


        }
        public void ResultPages(List<Table> tables)
        {
            try
            {
                // Kullanıcıdan kaydedilecek dosya yolu al
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Office Open WORD Belgesi (.docx)|*.docx";
                saveFileDialog1.Title = "Belgeyi Kaydet";
                DialogResult result = saveFileDialog1.ShowDialog();

                // Kullanıcı bir konum seçtiyse devam et
                if (result == DialogResult.OK && !string.IsNullOrEmpty(saveFileDialog1.FileName))
                {
                    string originalFilePath = "C:\\Users\\ABDURRAHMAN\\Desktop\\sertifikaB.docx";
                    string copyFilePath = saveFileDialog1.FileName;

                    // Belgeyi kopyala
                    File.Copy(originalFilePath, copyFilePath, true);

                    // Değiştirilecek değerleri ve karşılık gelecek yeni değerleri içeren sözlükler
                    Dictionary<string, List<Table>> replacements = new Dictionary<string, List<Table>>
            {
                { "ASTABLE", tables } // Tüm tabloları ekliyoruz
            };
                    Table table1 = new Table();
                    table1 = DeviceTable("Güç Algılayıcı", "Hewlett Packard", "8481A", "10Mhz-80Mhz");
                    Table table2 = new Table();
                    table2 = UsedDevice();

                    Dictionary<string, Table> replacementsTable = new Dictionary<string, Table>
        {

            { "UsedDevice", table2 },
            { "DeviceTable", table1 },
        };


                    Dictionary<string, string> replacements2 = new Dictionary<string, string>
            {
                { "AS1", "TUBITAK UME" },
                { "AS2", "2023.001" },
                { "AS3", "Güç Algılayıcı" },
                { "AS4", "Hewlett Packard" },
                { "AS5", "8481A" },
                { "AS6", "3318A97557" },
                { "KYY", "TUBITAK UME" },
                { "DOROD", DateTime.UtcNow.ToString("dd/MM/yyyy") },
                { "KYP", "Contrary to popular belief, Lorem Ipsum is not simply random text. It has roots in a piece of classical Latin literature from 45 BC, making it over 2000 years old. Richard McClintock, a Latin professor at Hampden-Sydney College in Virginia, looked up one of the more obscure Latin words, consectetur, from a Lorem Ipsum passage, and going through the cites of the word in classical literature, discovered the undoubtable source. Lorem Ipsum comes from sections 1.10.32 and 1.10.33 of \"de Finibus Bonorum et Malorum\" (The Extremes of Good and Evil) by Cicero, written in 45 BC. This book is a treatise on the theory of ethics, very popular during the Renaissance. The first line of Lorem Ipsum, \"Lorem ipsum dolor sit amet..\", comes from a line in section 1.10.32.\r\n\r\nThe standard chunk of Lorem Ipsum used since the 1500s is reproduced below for those interested. Sections 1.10.32 and 1.10.33 from \"de Finibus Bonorum et Malorum\" by Cicero are also reproduced in their exact original form, accompanied by English versions from the 1914 translation by H. Rackham." },
                { "AS7", DateTime.UtcNow.ToString("dd/MM/yyyy") },
                { "ASDATE", DateTime.UtcNow.ToString("dd/MM/yyyy") }
                // İhtiyacınıza göre buraya diğer değişikliklerinizi ekleyebilirsiniz
            };

                    // Belgeyi aç
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(copyFilePath, true))
                    {
                        var body = doc.MainDocumentPart.Document.Body;

                        // Tüm Text elementlerini gez
                        foreach (var text in body.Descendants<Text>().ToList())
                        {
                            // Her bir değiştirme işlemi için kontrol et
                            foreach (var replacement in replacements)
                            {
                                // Text içerisinde değiştirme anahtarını ara
                                if (text.Text.Contains(replacement.Key))
                                {
                                    // Eğer değiştirme değeri Table listesi ise
                                    if (replacement.Value is List<Table>)
                                    {
                                        var parent = text.Parent;

                                        // Tabloları belgeye ekleyin
                                        for (int i = replacement.Value.Count - 1; i >= 0; i--)
                                        {
                                            var table = replacement.Value[i];
                                            if (table != null)
                                            {
                                                // Tabloyu eklemek için mevcut paragrafın ebeveynini kullan
                                                parent.InsertAfterSelf(table.CloneNode(true));

                                                // Her bir tablodan sonra bir paragraf ekleyin
                                                parent.InsertAfterSelf(new Paragraph());
                                                parent.InsertAfterSelf(new Paragraph());
                                                parent.InsertAfterSelf(new Paragraph());
                                            }
                                        }

                                        // Orijinal text elementini kaldır
                                        text.Remove();
                                    }
                                }

                            }

                            // Metin parçalarını değiştir
                            foreach (var replacement2 in replacements2)
                            {
                                text.Text = text.Text.Replace(replacement2.Key, replacement2.Value);
                            }
                            foreach (var replacement in replacementsTable)
                            {
                                // Text içerisinde değiştirme anahtarını ara
                                if (text.Text.Contains(replacement.Key))
                                {
                                    var parent = text.Parent;

                                    // Tabloyu belgeye ekleyin
                                    parent.InsertAfterSelf(replacement.Value.CloneNode(true));

                                    // Her bir tablodan sonra bir paragraf ekleyin
                                    parent.InsertAfterSelf(new Paragraph());


                                    // Orijinal text elementini kaldır
                                    text.Remove();
                                }
                            }
                        }


                        // Belgeyi kaydet
                        doc.Save();
                    }

                    MessageBox.Show("İşlem Tamamlandı");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }

    }
}





