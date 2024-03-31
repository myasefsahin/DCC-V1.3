using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
     class RF_Difference_WordTable
    {
        public Table RF_Diff_Table(int sayac,string tableName,ArrayList col1, ArrayList col2, ArrayList col3, ArrayList col4, ArrayList col5, ArrayList col6, ArrayList col7,string Col1, string Col2, string Col3, string Col4, string Col5, string Col6, string Col7)
        {
            // Tablo özellikleri
            int tableWidth = 10000; // Sayfa genişliğine sığacak şekilde ayarlanmalıdır (örneğin, Word'de varsayılan 10000 DXA)
            int cellWidth = tableWidth / 7; // Yedi sütun var, bu yüzden hücre genişliğini buna göre ayarlayın

            // Tablo başlığı


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
                     new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },new FontSize() { Val="22"}),
                     new Text("Tablo "+sayac+". "+tableName)
                 )
             )
         )
     );
            TableRow headerRow = new TableRow(
                  new TableRowProperties(
                new TableHeader()
            ),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(),new FontSize() { Val = "22" }),
                        new Text(Col1)))),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(),new FontSize() { Val = "22" }),
                        new Text(Col2)))),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text(Col3)))),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold()), new FontSize() { Val = "22" },
                        new Text(Col4)))),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text(Col5)))),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text(Col6)))),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text(Col7))))
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
                headerRow1,headerRow
            );

            // Veri satırlarını ekle
            for (int i = 0; i < col1.Count; i++)
            {
                TableRow dataRow = new TableRow(
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(col1[i].ToString())))),
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(col2[i].ToString())))),
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(col3[i].ToString())))),
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(col4[i].ToString())))),
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(col5[i].ToString())))),
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(col6[i].ToString())))),
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(col7[i].ToString()))))
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
