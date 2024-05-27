using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    class CF_WordTable
    {
     
        //Calibration Factor ölçüm tipinin Excel üzerindeki tabloların tablo yapılarına göre oluşturulma kodu, 
        //Bu method seçilen dosyanın adını Certificate form üzerinde tutulan sayacı, ve CF_DataWord classında çekilen ve formatlanan verileri parametre olarak alır. 
        public Table CF_CreateCF(string filename,int sayac,string tableName,ArrayList Freq, ArrayList CF, ArrayList CF_Unc)
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
                   new GridSpan() { Val = 3 } // 3 hücreyi birleştir
               ),
               new Paragraph(
                   new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center } // Paragrafı ortala
                   ),
                   new Run(
                       //Tablo başlığındaki verinin yazıldığı kısım 
                       new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() {Val="22" }),
                       new Text("Tablo" + sayac + ". " + tableName + "-" + filename)
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
                        new Text("CF")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("CF Uncertainty"))))
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
                            new Text(CF[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(CF_Unc[i].ToString()))))
                );

                table.Append(dataRow);
            }

            return table;
        }


        public Table CF_CreateReelImg(string filename, int sayac,string tableName,ArrayList Freq, ArrayList Reel, ArrayList ReelUnc, ArrayList Imag, ArrayList ImagUnc, ArrayList YK, ArrayList YK_Unc)
        {
            // Tablo özellikleri
            int tableWidth = 10000; // Sayfa genişliğine sığacak şekilde ayarlanmalıdır (örneğin, Word'de varsayılan 10000 DXA)
            int cellWidth = tableWidth / 7; // Yedi sütun var, bu yüzden hücre genişliğini buna göre ayarlayın

        TableRow headerRow1 = new TableRow(
        new TableRowProperties(
           new TableHeader()
        ),
        new TableCell(
           new TableCellProperties(
               new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() },
               new GridSpan() { Val = 7 } 
           ),
           new Paragraph(
               new ParagraphProperties(
                   new Justification() { Val = JustificationValues.Center } // Paragrafı ortala
               ),
               new Run(
                   new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },  new FontSize() { Val = "22" }),
                       new Text("Tablo" + sayac + ". " + tableName + "-" + filename)
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
                        new Text("Frekans (GHz)")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Gerçel Bileşen (x)")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Gerçel Bileşen Belirsizliği")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Sanal Bileşen (y)")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Sanal Bileşen Belirsizliği")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Yansıma Katsayısı")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Yansıma Katsayısı Belirsizliği"))))
            );

            // Tablo oluştur
            Table table = new Table(
                new TableProperties(
                    new TableStyle() { Val = "TableGrid" },
                    new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = tableWidth.ToString() },
                    new TableLook() { Val = "04A0" },
                    // Kenar çizgileri ekle
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
                            new Text(ImagUnc[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(YK[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(YK_Unc[i].ToString()))))
                );

                // Hücrelerin kenar çizgilerini ayarla
                foreach (TableCell cell in dataRow.Elements<TableCell>())
                {
                    TableCellProperties cellProperties = cell.GetFirstChild<TableCellProperties>();
                    if (cellProperties == null)
                    {
                        cellProperties = new TableCellProperties();
                        cell.AppendChild(cellProperties);
                    }

                    cellProperties.Append(new TableCellBorders(
                        new TopBorder() { Val = BorderValues.Single, Size = 9 },
                        new BottomBorder() { Val = BorderValues.Single, Size = 9 },
                        new LeftBorder() { Val = BorderValues.Single, Size = 9 },
                        new RightBorder() { Val = BorderValues.Single, Size = 9 }
                    ));
                }

                table.Append(dataRow);
            }

            return table;
        }


    }
}
