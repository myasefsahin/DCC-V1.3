using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    class RF_Gain_WordTable
    {
        public Table RF_Gain_Table(int sayac,string tableName,ArrayList col1, ArrayList col2, ArrayList col3, string Col1, string Col2, string Col3)
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
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },new FontSize() { Val="22"}),
                        new Text("Tablo "+sayac+". "+tableName)
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
                        new Text(Col1)))),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text(Col2)))),
                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text(Col3))))
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

            for (int i = 0; i < col1.Count; i++)
            {
                TableRow dataRow = new TableRow(
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(col1[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(col2[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(col3[i].ToString()))))
                );

                table.Append(dataRow);
            }

            return table;
        }
    }
}
