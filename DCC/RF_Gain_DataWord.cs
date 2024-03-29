using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    class RF_Gain_DataWord
    {
        List<string> columnName = new List<string>(104) {
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
                "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ","AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
                "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ","BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT","BU", "BV", "BW", "BX", "BY", "BZ",
                "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ","CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT","CU", "CV", "CW", "CX", "CY", "CZ" };


        public void main(string ExcelDosyaYolu, string pageName, int satır, string sütun)
        {

            int harfIndex = columnName.IndexOf(sütun);
            using (var package = new ExcelPackage(new FileInfo(ExcelDosyaYolu)))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets[pageName];
                int rowCount = worksheet.Dimension.End.Row;
                string[] cellValue = new string[rowCount];


                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[sütun + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        RFG_T1_Frekans.Add(cellValue[i - satır]);
                    }
                }


                for (int i = satır; i < RFG_T1_Frekans.Count + satır; i++)
                {

                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 1] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 2] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    RFG_T1_GirisGucu.Add(formattedEntity.measurent);
                    RFG_T1_Belirsizlik.Add(formattedEntity.uncertainty);

                  

                }


                for (int i = satır; i <= rowCount; i++)
               {
                     cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 4] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        RFG_T2_EnBuyukKazanc.Add(cellValue[i - satır]);
                    }
                }

                for (int i = satır; i < RFG_T2_EnBuyukKazanc.Count + satır; i++)
                {
                    RFG_T2_EnKucukKazanc.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 5] + i].Value));
                    RFG_T2_Flatness.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 6] + i].Value));
                }


                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 8] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        RFG_T3_Nom_Giris_Gucu.Add(cellValue[i - satır]);
                    }
                }


                for (int i = satır; i < RFG_T3_Nom_Giris_Gucu.Count + satır; i++)
                {

                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 9] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 10] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    RFG_T3_Kazanc.Add(formattedEntity1.measurent);
                    RFG_T3_Belirsizlik.Add(formattedEntity1.uncertainty);



                }


                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 12] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        RFG_T4_Nom_Giris_Gucu.Add(cellValue[i - satır]);
                    }
                }


                for (int i = satır; i < RFG_T4_Nom_Giris_Gucu.Count + satır; i++)
                {
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 13] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 14] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    RFG_T4_Kazanc.Add(formattedEntity1.measurent);
                    RFG_T4_Belirsizlik.Add(formattedEntity1.uncertainty);



                }
            }
        }

        public ArrayList RFG_T1_Frekans { get; set; }
        public ArrayList RFG_T1_GirisGucu { get; set; }
        public ArrayList RFG_T1_Belirsizlik { get; set; }
 
        public ArrayList RFG_T2_EnBuyukKazanc { get; set; }
        public ArrayList RFG_T2_EnKucukKazanc { get; set; }
        public ArrayList RFG_T2_Flatness { get; set; }


        public ArrayList RFG_T3_Nom_Giris_Gucu { get; set; }
        public ArrayList RFG_T3_Kazanc { get; set; }
        public ArrayList RFG_T3_Belirsizlik { get; set; }

        public ArrayList RFG_T4_Nom_Giris_Gucu { get; set; }
        public ArrayList RFG_T4_Kazanc { get; set; }
        public ArrayList RFG_T4_Belirsizlik { get; set; }


    public RF_Gain_DataWord()
    {
        RFG_T1_Frekans = new ArrayList();
        RFG_T1_GirisGucu = new ArrayList();
        RFG_T1_Belirsizlik = new ArrayList();

        RFG_T2_EnBuyukKazanc = new ArrayList();
        RFG_T2_EnKucukKazanc = new ArrayList();
        RFG_T2_Flatness = new ArrayList();

        RFG_T3_Nom_Giris_Gucu = new ArrayList();
        RFG_T3_Kazanc = new ArrayList();
        RFG_T3_Belirsizlik = new ArrayList();

        RFG_T4_Nom_Giris_Gucu = new ArrayList();
        RFG_T4_Kazanc = new ArrayList();
        RFG_T4_Belirsizlik = new ArrayList();


    }
  }
}
