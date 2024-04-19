using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
     class Noise_WordTable {

        public Table CreateENR(int sayac, string tableName, ArrayList Freq, ArrayList ENR, ArrayList ENRUnc)
        {
                    int tableWidth = 10000;
            int cellWidth = tableWidth / 3;


          TableRow headerRow1 = new TableRow(
            new TableRowProperties(
                new TableHeader()
            ),
            new TableCell(
                new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() },
                    new GridSpan() { Val = 3 } // 3 hücreyi birleştir
                ),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center } // Paragrafı ortala
                    ),
                    new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val="22"}),
                        new Text("Tablo " + sayac + ". " + tableName)
                    )
                )
            )
        );

            // Tablo başlığı
            TableRow headerRow = new TableRow(
                new TableProperties(
                     new TableHeader()
                     ),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Frequency (GHz)")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("ENR")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("ENR Uncertainty"))))
            );

            // Tablo oluştur
            Table table = new Table(
                new TableProperties(
                    new TableStyle() { Val = "TableGrid" },
                    new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = tableWidth.ToString() },
                    new TableLook() { Val = "04A0" },
                       new TableBorders(
                        new TopBorder() { Val = BorderValues.Single, Size = 9 },
                        new BottomBorder() { Val = BorderValues.Single, Size = 9 },
                        new LeftBorder() { Val = BorderValues.Single, Size = 9 },
                        new RightBorder() { Val = BorderValues.Single, Size = 9 },
                        new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 9 },
                        new InsideVerticalBorder() { Val = BorderValues.Single, Size = 9 }
                    )
                ),
               headerRow1, headerRow
            );

            for (int i = 0; i < Freq.Count; i++)
            {
                TableRow dataRow = new TableRow(
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(Freq[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(ENR[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(ENRUnc[i].ToString()))))
                );

                table.Append(dataRow);
            }

            return table;
        }

        public Table Create_DC_ON_OFF(int sayac, string tableName, ArrayList frekans, ArrayList RC, ArrayList UstLimit, ArrayList RCUnc, ArrayList Phase, ArrayList PhaseUnc)
        {
            // Table properties and some other specs
            int tableWidth = 10000;
            int cellWidth = tableWidth / 3;


            TableRow headerRow1 = new TableRow(
              new TableRowProperties(
                  new TableHeader()
              ),
              new TableCell(
                  new TableCellProperties(
                      new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() },
                      new GridSpan() { Val = 6 } // 3 hücreyi birleştir
                  ),
                  new Paragraph(
                      new ParagraphProperties(
                          new Justification() { Val = JustificationValues.Center } // Paragrafı ortala
                      ),
                      new Run(
                          new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                          new Text("Tablo " + sayac + ". " + tableName)
                      )
                  )
              )
          );

            // Tablo başlığı
            TableRow headerRow = new TableRow(
                new TableProperties(
                     new TableHeader()
                     ),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("frequency (ghz)")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Reflection Coefficient Linear Amplitude")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Reflection coefficient upper limit")))),

                 new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Reflection coefficient linear amplitude uncertainty")))),

                  new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Reflection Coefficient Phase")))),

                   new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Reflection Coefficient Phase Unc"))))
            );

            // Tablo oluştur
            Table table = new Table(
                new TableProperties(
                    new TableStyle() { Val = "TableGrid" },
                    new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = tableWidth.ToString() },
                    new TableLook() { Val = "04A0" },
                       new TableBorders(
                        new TopBorder() { Val = BorderValues.Single, Size = 9 },
                        new BottomBorder() { Val = BorderValues.Single, Size = 9 },
                        new LeftBorder() { Val = BorderValues.Single, Size = 9 },
                        new RightBorder() { Val = BorderValues.Single, Size = 9 },
                        new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 9 },
                        new InsideVerticalBorder() { Val = BorderValues.Single, Size = 9 }
                    )
                ),
               headerRow1, headerRow
            );

            for (int i = 0; i < frekans.Count; i++)
            {
                TableRow dataRow = new TableRow(
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(frekans[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(RC[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(UstLimit[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(RCUnc[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(Phase[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(PhaseUnc[i].ToString()))))
                );

                table.Append(dataRow);
            }

            return table;
        }

    }
}
