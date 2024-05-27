using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections;
namespace DCC
{
    class SP_WordTable
    {
        #region Reel & Imaginary
        public Table CreateReelImg(string filename,int sayac,string tableName,ArrayList Freq, ArrayList Reel, ArrayList ReelUnc, ArrayList Imag, ArrayList ImagUnc)
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
                      new GridSpan() { Val = 5 } 
                  ),
                  new Paragraph(
                      new ParagraphProperties(
                          new Justification() { Val = JustificationValues.Center } // Paragrafı ortala
                      ),
                      new Run(
                          new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val="22"}),
                       new Text("Tablo" + sayac + ". " + tableName + " " + "(" + filename + ")")
                      )
                  )
              )
          );

            
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
                        new Text("Reel Component (x)")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Reel Component Uncertainty")))),

                 new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Imaginary Component (y)")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Imaginary Component Uncertainty"))))
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
                            new Text(Reel[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(ReelUnc[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(Imag[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(ImagUnc[i].ToString()))))
                );

                table.Append(dataRow);
            }

            return table;
        }

        #endregion

        #region Linmag & Phase
        public Table CreateLinPhase(string filename, int sayac, string tableName, ArrayList Freq, ArrayList LinMag, ArrayList LinMagUnc, ArrayList Phase, ArrayList PhaseUnc)
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
                      new GridSpan() { Val = 5 } // 3 hücreyi birleştir
                  ),
                  new Paragraph(
                      new ParagraphProperties(
                          new Justification() { Val = JustificationValues.Center } // Paragrafı ortala
                      ),
                      new Run(
                          new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                       new Text("Tablo" + sayac + ". " + tableName + " " + "(" + filename + ")")
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
                        new Text("Linear Magnitude")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Linear Magnitude Uncertainty")))),

                 new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val="22"}),
                        new Text("Phase")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Phase Uncertainty"))))
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
                            new Text(LinMag[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(LinMagUnc[i].ToString())))),

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

        #endregion

        #region Logmag & Phase
        public Table CreateLogPhase(string filename, int sayac, string tableName, ArrayList Freq, ArrayList LogMag, ArrayList LogMagUnc, ArrayList LogPhase, ArrayList LogPhaseUnc)
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
                      new GridSpan() { Val = 5 }
                  ),
                  new Paragraph(
                      new ParagraphProperties(
                          new Justification() { Val = JustificationValues.Center } // Paragrafı ortala
                      ),
                      new Run(
                          new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                       new Text("Tablo" + sayac + ". " + tableName + " " + "(" + filename + ")")
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
                        new Text("Logarithmic Magnitude")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Logarithmic Magnitude Uncertainty")))),

                 new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Phase")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Phase Uncertainty"))))
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
                            new Text(LogMag[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(LogMagUnc[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(LogPhase[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(LogPhaseUnc[i].ToString()))))
                );

                table.Append(dataRow);
            }

            return table;   
        }

        #endregion

        #region SWR
        public Table CreateSWR(string filename,int sayac, string tableName, ArrayList Freq, ArrayList Swr, ArrayList SwrUnc)
        {
            // Tablo özellikleri
            int tableWidth = 10000; // Sayfa genişliğine sığacak şekilde ayarlanmalıdır (örneğin, Word'de varsayılan 10000 DXA)
            int cellWidth = tableWidth / 3; // Üç sütun var, bu yüzden hücre genişliğini buna göre ayarlayın

            TableRow headerRow1 = new TableRow(
           new TableRowProperties(
               new TableHeader()
           ),
           new TableCell(
               new TableCellProperties(
                   new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() },
                   new GridSpan() { Val = 3 } 
               ),
               new Paragraph(
                   new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center } // Paragrafı ortala
                   ),
                   new Run(
                       new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                       new Text("Tablo" + sayac + ". " + tableName + " " + "(" + filename + ")")
                   )
               )
           )
       );

            // Tablo başlığı
            TableRow headerRow = new TableRow(
                  new TableRowProperties(
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
                        new Text("SWR")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("SWR Uncertainty"))))
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

            // Veri satırlarını ekle
            for (int i = 0; i < Freq.Count; i++)
            {
                TableRow dataRow = new TableRow(
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(Freq[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(Swr[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(SwrUnc[i].ToString()))))
                );

                table.Append(dataRow);
            }

            return table;
        }
        #endregion

        
    }


}

