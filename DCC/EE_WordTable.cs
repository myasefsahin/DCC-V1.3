using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    class EE_WordTable
    {
        public Table EECreateReelImg(ArrayList Freq, ArrayList Reel, ArrayList ReelUnc, ArrayList Imag, ArrayList ImagUnc)
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


        public Table CreateEE(ArrayList Freq, ArrayList EE, ArrayList EE_Unc)
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
                            new Text("EE"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(new RunProperties(new Italic()), new Text("EE")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "4430" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("EE Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("EE Uncertainty"))))
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
                            new Text(EE[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(EE_Unc[i].ToString()))))
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

        public Table CreateRHO(ArrayList Freq, ArrayList Rho_Lin, ArrayList RhoUnc)
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
                            new Text("RHO lin"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(new RunProperties(new Italic()), new Text("RHO lin")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "4430" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("RHO Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("RHO Uncertainty"))))
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
                            new Text(Rho_Lin[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(RhoUnc[i].ToString()))))
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
        public Table CreateCF(ArrayList Freq, ArrayList CF, ArrayList CF_Unc)
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
                            new Text("CF"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(new RunProperties(new Italic()), new Text("CF")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "4430" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("CF Belirsizliği"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("CF Uncertainty"))))
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
                            new Text(CF[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(CF_Unc[i].ToString()))))
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
