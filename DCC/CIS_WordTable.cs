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
        public Table Create_Z_Position(int sayac,string tableName,ArrayList Olcum_Adim, ArrayList z_position, ArrayList z_position_Unc, ArrayList Icod, ArrayList Icod_Unc, ArrayList Ocid, ArrayList Ocid_Unc)
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
                    new GridSpan() { Val = 7 } // 3 hücreyi birleştir
                ),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center } // Paragrafı ortala
                    ),
                    new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },new FontSize() { Val="22" }),
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
                        new Text("Ölçüm Adımları(Z-Position)")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("Z-Position")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("U(Z-Position)")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("ICOD")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("U(ICOD)")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("OCID")))),

                new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                    new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                        new Run(
                        new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new Bold(), new FontSize() { Val = "22" }),
                        new Text("U(OCID)"))))
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

            for (int i = 0; i < Olcum_Adim.Count; i++)
            {
                TableRow dataRow = new TableRow(
                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(Olcum_Adim[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(z_position[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(z_position_Unc[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(Icod[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(Icod_Unc[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(Ocid[i].ToString())))),

                    new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidth.ToString() }),
                        new Paragraph(new ParagraphProperties(
                       new Justification() { Val = JustificationValues.Center }),
                            new Run(
                            new RunProperties(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize() { Val = "22" }),
                            new Text(Ocid_Unc[i].ToString()))))
                   
                );

                table.Append(dataRow);
            }

            return table;

        }
    }
}
