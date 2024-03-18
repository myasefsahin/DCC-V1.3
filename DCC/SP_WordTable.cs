using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections;
namespace DCC
{
    class SP_WordTable
    {
        #region Reel & Imaginary
        public Table CreateReelImg(ArrayList Freq, ArrayList Reel, ArrayList ReelUnc, ArrayList Imag, ArrayList ImagUnc)
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
                            new Text("Gerçel Bileşen (x)"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Reel Component (x)")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Gerçel Bileşen Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Reel Component Uncertainty")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Sanal Bileşen (y)"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Imaginary Component (y)")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Sanal Bileşen Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Imaginary Component Uncertainty"))))
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

                TableRow row = new TableRow(

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Freq[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Reel[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(ReelUnc[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Imag[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(ImagUnc[i].ToString()))))
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

        #region Linmag & Phase
        public Table CreateLinPhase(ArrayList Freq, ArrayList LinMag, ArrayList LinMagUnc, ArrayList Phase, ArrayList PhaseUnc)
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
                            new Text("Doğrusal Büyüklük"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Linear Magnitude")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Doğrusal Büyüklük Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Linear Magnitude Uncertainty")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Faz"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Phase")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Faz Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Phase Uncertainty"))))
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

                //double linData = (double)System.Math.Round(Convert.ToDouble(LinMag[i]), 4);
                //double linuncData = (double)System.Math.Round(Convert.ToDouble(LinMagUnc[i]), 4);
                //double fazData = (double)System.Math.Round(Convert.ToDouble(Phase[i]), 4);
                //double fazuncData = (double)System.Math.Round(Convert.ToDouble(PhaseUnc[i]), 4);

                TableRow row = new TableRow(

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Freq[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(LinMag[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(LinMagUnc[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Phase[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(PhaseUnc[i].ToString()))))
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

        #region Logmag & Phase
        public Table CreateLogPhase(ArrayList Freq, ArrayList LogMag, ArrayList LogMagUnc, ArrayList LogPhase, ArrayList LogPhaseUnc)
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
                            new Text("Logaritmik Büyüklük"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Logarithmic Magnitude")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Logaritmik Büyüklük Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Logarithmic Magnitude Uncertainty")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Faz"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Phase")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Faz Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Phase Uncertainty"))))
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
                            new Text(LogMag[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(LogMagUnc[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(LogPhase[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(LogPhaseUnc[i].ToString()))))
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

        #region SWR
        public Table CreateSWR(ArrayList Freq, ArrayList Swr, ArrayList SwrUnc)
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
                            new Text("SWR"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(new RunProperties(new Italic()), new Text("SWR")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "4430" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("SWR Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("SWR Uncertainty"))))
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
                            new Text(Swr[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(SwrUnc[i].ToString()))))
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
        public Table CoverPage(string Order, string Device, string Serial)
        {
            Table table = new Table(
                new TableProperties(
                new TableStyle() { Val = "TableGrid" },
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "100%" },
                new TableLook() { Val = "04A0" }
                )
            );

            bool isFirstRow = true;

            int i = 1;
            string FullText;
            // Create a paragraph for each of the Order, Device, and Serial strings
            foreach (string text in new[] { Order, Device, Serial })
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
                    FullText = "Order Number: " + Order;
                }
                else if (i == 2)
                {
                    FullText = "Model Name: " + Device;
                }
                else if (i == 3)
                {
                    FullText = "Serial Number: " + Serial;
                }
                else
                {
                    FullText = "NULL";
                }
                // Create a paragraph for the current text
                Paragraph paragraph = new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center }
                    ),
                    new Run(
                        new RunProperties(
                            new Bold(),
                            new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                            new FontSize() { Val = "32" } // 16 point font size (twice the font size for Word)
                        ),
                        new Text(FullText)
                    )
                );

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
    }


}

