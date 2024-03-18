using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    class CIS_WordTable
    {
        public Table Create_Z_Position(ArrayList Olcum_Adim, ArrayList z_position, ArrayList z_position_Unc, ArrayList Icod, ArrayList Icod_Unc, ArrayList Ocid, ArrayList Ocid_Unc)
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
                            new Text("Ölçüm Adımları(Z-Position)"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Measurement Steps (Z-Position)")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("Z-Position"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("Z-Position")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("U(Z-Position)"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("U(Z-Position)")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("ICOD"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("ICOD")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("U(ICOD)"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("U(ICOD)")))),

                     new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("OCID"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("OCID")))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()),
                            new Text("U(OCID)"))),
                        new Break() { Type = BreakValues.TextWrapping },
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Italic()),
                            new Text("U(OCID)"))))
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
            for (int i = 0; i < Olcum_Adim.Count; i++)
            {

                TableRow row = new TableRow(

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1488" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Olcum_Adim[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(z_position[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(z_position_Unc[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Icod[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Icod_Unc[i].ToString())))),

                   new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Ocid[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2215" }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }),
                            new Text(Ocid_Unc[i].ToString()))))
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
