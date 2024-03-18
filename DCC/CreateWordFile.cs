using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO; // System.IO namespace'i eklenmeli
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    public class CreateWordFile
    {
        // Sabit olarak şablon dosyası adını tanımlayın
        public string wordPath;

        public void CreateWord(string fileName, string filePath)
        {
            wordPath = filePath + "\\" + fileName + ".docx";

            using (WordprocessingDocument doc = WordprocessingDocument.Create(wordPath, WordprocessingDocumentType.Document))
            {

                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // Set page measurements
                SetPageMeasurements(mainPart.Document.Body);
            }
        }
        public void ProcessWord(List<Table> tables, List<string> header)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(wordPath, true))
            {
                Body body = doc.MainDocumentPart.Document.Body;

                int i = 1;
                // Tables goes here
                foreach (Table table in tables)
                {
                    Paragraph pageBreak = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));

                    Paragraph titleParagraph = new Paragraph(
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Center }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                                new Bold(),
                                new Italic()
                            ),
                            new Text(header[i - 1])
                        ),
                        new Run(
                            new Break() { Type = BreakValues.TextWrapping }
                        )
                    );

                    body.Append(titleParagraph);
                    body.Append(table);

                    if (i < tables.Count)
                        body.Append(pageBreak);

                    i++;
                }
            }
        }
        public static void SetPageMeasurements(Body body)
        {
            SectionProperties sectionProps = new SectionProperties(
                new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U, Orient = PageOrientationValues.Portrait },
                new PageMargin()
                {
                    Top = 624,
                    Right = 794U,
                    Bottom = 369,
                    Left = 907U,
                    Header = 340U,
                    Footer = 284U,
                    Gutter = 0U
                }
            );

            body.Append(sectionProps);
        }
    }
}
