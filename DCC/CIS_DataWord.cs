
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    public class CIS_DataWord
    {
        List<string> columnName = new List<string>(104) {
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
                "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ","AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
                "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ","BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT","BU", "BV", "BW", "BX", "BY", "BZ",
                "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ","CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT","CU", "CV", "CW", "CX", "CY", "CZ" };
        public void main(string ExcelDosyaYolu, string pageName,int satır, string sütun)
        {
            int harfIndex = columnName.IndexOf(sütun);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Excel Verileri

            using (var package = new ExcelPackage(new FileInfo(ExcelDosyaYolu)))
            {


                // Excel'in sabit olan sonuç sayfasındaki verilere göre hazırlanmıştır. 
                ExcelWorksheet worksheet = package.Workbook.Worksheets[pageName];

                int rowCount = worksheet.Dimension.Rows;
                //int satirSayisi = rowCount - 6;

                string[] cellValue = new string[rowCount];


                // Frekans değerlerinin çekimi
                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[sütun + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        CIS_Olcum_Adım.Add(cellValue[i - satır]);
                    }
                }

                for (int i = satır; i < CIS_Olcum_Adım.Count+ satır;  i++)
                {

                    //S11 değerleri için 
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();



                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex+1] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 2] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    CIS_ZP.Add(formattedEntity.measurent);
                    CIS_ZP_Unc.Add(formattedEntity.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 3] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 4] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    CIS_ICOD.Add(formattedEntity1.measurent);
                    CIS_ICOD_Unc.Add(formattedEntity1.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 5] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 6] + i].Value);
                    CalculateEntity formattedEntity2 = NumberFormatter.deneme(calculateEntity);
                    CIS_OCID.Add(formattedEntity.measurent);
                    CIS_OCID_Unc.Add(formattedEntity2.uncertainty);





                }

            }
        }
        public ArrayList CIS_Olcum_Adım { get; set; }



        public ArrayList CIS_ZP { get; set; }
        public ArrayList CIS_ZP_Unc { get; set; }
        public ArrayList CIS_ICOD { get; set; }
        public ArrayList CIS_ICOD_Unc { get; set; }
        public ArrayList CIS_OCID { get; set; }
        public ArrayList CIS_OCID_Unc { get; set; }





        public CIS_DataWord()
        {
            CIS_Olcum_Adım = new ArrayList();

            CIS_ZP = new ArrayList();
            CIS_ZP_Unc = new ArrayList();
            CIS_ICOD = new ArrayList();
            CIS_ICOD_Unc = new ArrayList();
            CIS_OCID = new ArrayList();
            CIS_OCID_Unc = new ArrayList();

        }

        public void ClearData()
        {
            CIS_Olcum_Adım.Clear();
            CIS_ZP.Clear();
            CIS_ZP_Unc.Clear();
            CIS_ICOD.Clear();
            CIS_ICOD_Unc.Clear();
            CIS_OCID.Clear();
            CIS_OCID_Unc.Clear();

        }
    }


}
