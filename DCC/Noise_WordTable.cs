using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
     class Noise_WordTable {

        public Table CreateENR(ArrayList Freq, ArrayList ENR, ArrayList ENRUnc)
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
                            new Text("ENR"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(new RunProperties(
                            new Italic()), 
                        new Text("ENR")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "4430" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("ENR Uncertainty"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("ENR Uncertainty"))))
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
                            new Text(ENR[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(ENRUnc[i].ToString()))))
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

        public Table Create_DC_ON_OFF(ArrayList frekans, ArrayList RC, ArrayList UstLimit, ArrayList RCUnc, ArrayList Phase, ArrayList PhaseUnc, ArrayList Kontrol)
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
                            new Text("frekans(ghz)"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("frequency (ghz)")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Yansıma Katsayısı Lineer Genliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Reflection Coefficient Linear Amplitude")))),

                     new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Yansıma katsayısı üst limiti"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Reflection coefficient upper limit")))),


                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("yansıma katsayısı lineer genlik belirsizliği "))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Reflection coefficient linear amplitude uncertainty")))),


                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Yansıma Katsayısı Fazı"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Reflection Coefficient Phase")))),

                      new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Yansıma Katsayısı Fazı belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Reflection Coefficient Phase Unc")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Kontrol"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Control"))))


                ),
                new TableJustification() { Val = TableRowAlignmentValues.Center }
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


            for (int i = 0; i < frekans.Count; i++)
            {

                TableRow row = new TableRow(

                     new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(frekans[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(RC[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(UstLimit[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(RCUnc[i].ToString())))),


                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Phase[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(PhaseUnc[i].ToString())))),

                   new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Kontrol[i].ToString()))))

               
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

    }
}
