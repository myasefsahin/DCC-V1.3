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
    class EE_DataWord
    {
        List<string> columnName = new List<string>(104) {
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
                "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ","AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
                "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ","BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT","BU", "BV", "BW", "BX", "BY", "BZ",
                "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ","CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT","CU", "CV", "CW", "CX", "CY", "CZ" }; 


        public void main(string ExcelDosyaYolu, string pageName, int satır, string sütun)
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
                        EE_ArrayFrekans.Add(cellValue[i - satır]);
                    }
                }


                // S-Parametre değerlerinin çekimi
                for (int i = satır; i < EE_ArrayFrekans.Count + satır; i++)
                {


                    //S11 değerleri için 
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 1] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 2] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    EE_ArrayEE.Add(formattedEntity.measurent);
                    EE_ArrayEEUnc.Add(formattedEntity.uncertainty);


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 4] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 5] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    EE_ArrayS11Reel.Add(formattedEntity1.measurent);
                    EE_ArrayS11ReelUnc.Add(formattedEntity1.uncertainty);
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 6] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 7] + i].Value);
                    CalculateEntity formattedEntity7 = NumberFormatter.deneme(calculateEntity);
                    EE_ArrayS11Complex.Add(formattedEntity7.measurent);
                    EE_ArrayS11ComplexUnc.Add(formattedEntity7.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 9] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 10] + i].Value);
                    CalculateEntity formattedEntity8 = NumberFormatter.deneme(calculateEntity);
                    EE_ArrayRhoLin.Add(formattedEntity8.measurent);
                    EE_ArrayRhoUnc.Add(formattedEntity8.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 12] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 13] + i].Value);
                    CalculateEntity formattedEntity13 = NumberFormatter.deneme(calculateEntity);
                    EE_ArrayCF.Add(formattedEntity13.measurent);
                    EE_ArrayCFUnc.Add(formattedEntity13.uncertainty);



                }
            }

        }

        public ArrayList EE_ArrayFrekans { get; set; }


        public ArrayList EE_ArrayEE { get; set; }
        public ArrayList EE_ArrayEEUnc { get; set; }

        public ArrayList EE_ArrayS11Reel { get; set; }
        public ArrayList EE_ArrayS11ReelUnc { get; set; }
        public ArrayList EE_ArrayS11Complex { get; set; }
        public ArrayList EE_ArrayS11ComplexUnc { get; set; }

        public ArrayList EE_ArrayRhoLin { get; set; }
        public ArrayList EE_ArrayRhoUnc { get; set; }

        public ArrayList EE_ArrayCF { get; set; }
        public ArrayList EE_ArrayCFUnc { get; set; }



        // Device Informations

        public string OrderNumber { get; set; }
        public string DeviceName { get; set; }
        public string SerialNumber { get; set; }


        public EE_DataWord()
        {
            EE_ArrayFrekans = new ArrayList();

            EE_ArrayEE = new ArrayList();
            EE_ArrayEEUnc = new ArrayList();

            EE_ArrayS11Reel = new ArrayList();
            EE_ArrayS11ReelUnc = new ArrayList();
            EE_ArrayS11Complex = new ArrayList();
            EE_ArrayS11ComplexUnc = new ArrayList();

            EE_ArrayRhoLin = new ArrayList();
            EE_ArrayRhoUnc = new ArrayList();

            EE_ArrayCF = new ArrayList();
            EE_ArrayCFUnc = new ArrayList();

        }

        public void ClearData()
        {
            EE_ArrayFrekans.Clear();
            EE_ArrayEE.Clear();
            EE_ArrayEEUnc.Clear();
            EE_ArrayS11Reel.Clear();
            EE_ArrayS11ReelUnc.Clear();
            EE_ArrayS11Complex.Clear();
            EE_ArrayS11ComplexUnc.Clear();
            EE_ArrayRhoLin.Clear();
            EE_ArrayRhoUnc.Clear();
            EE_ArrayCF.Clear();
            EE_ArrayCFUnc.Clear();

        }
    }
}
