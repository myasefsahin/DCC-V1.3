
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

    class CF_DataWord
    {
        public void main(string ExcelDosyaYolu, string pageName)
        {
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
                for (int i = 4; i <= rowCount; i++)
                {
                    cellValue[i - 4] = Convert.ToString(worksheet.Cells["B" + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - 4]))
                    {
                        CF_ArrayFrekans.Add(cellValue[i - 4]);
                    }
                }


                // CF Değerlerinin çekimi
                for (int i = 4; i < CF_ArrayFrekans.Count + 4; i++)
                {

                    //S11 değerleri için 
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells["C" + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells["D" + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    CF_Array.Add(formattedEntity.measurent);
                    CF_ArrayCFUnc.Add(formattedEntity.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells["H" + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells["I" + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    CF_ArrayReel.Add(formattedEntity1.measurent);
                    CF_ArrayReelUnc.Add(formattedEntity1.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells["J" + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells["K" + i].Value);
                    CalculateEntity formattedEntity2 = NumberFormatter.deneme(calculateEntity);
                    CF_ArrayComplex.Add(formattedEntity2.measurent);
                    CF_ArrayComplexUnc.Add(formattedEntity2.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells["L" + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells["M" + i].Value);
                    CalculateEntity formattedEntity3 = NumberFormatter.deneme(calculateEntity);
                    CF_YK.Add(formattedEntity3.measurent);
                    CF_YK_Unc.Add(formattedEntity3.uncertainty);





                }
            }

        }

        public ArrayList CF_ArrayFrekans { get; set; }


        public ArrayList CF_Array { get; set; }
        public ArrayList CF_ArrayCFUnc { get; set; }

        public ArrayList CF_ArrayReel { get; set; }
        public ArrayList CF_ArrayReelUnc { get; set; }
        public ArrayList CF_ArrayComplex { get; set; }
        public ArrayList CF_ArrayComplexUnc { get; set; }
        public ArrayList CF_YK { get; set; }
        public ArrayList CF_YK_Unc { get; set; }




        // Device Informations

        public string OrderNumber { get; set; }
        public string DeviceName { get; set; }
        public string SerialNumber { get; set; }


        public CF_DataWord()
        {
            CF_ArrayFrekans = new ArrayList();

            CF_Array = new ArrayList();
            CF_ArrayCFUnc = new ArrayList();

            CF_ArrayReel = new ArrayList();
            CF_ArrayReelUnc = new ArrayList();
            CF_ArrayComplex = new ArrayList();
            CF_ArrayComplexUnc = new ArrayList();

            CF_YK = new ArrayList();
            CF_YK_Unc = new ArrayList();


        }

        public void ClearData()
        {
            CF_ArrayFrekans.Clear();
            CF_Array.Clear();
            CF_ArrayCFUnc.Clear();
            CF_ArrayReel.Clear();
            CF_ArrayReelUnc.Clear();
            CF_ArrayComplex.Clear();
            CF_ArrayComplexUnc.Clear();
            CF_YK.Clear();
            CF_YK_Unc.Clear();


        }
    }
}
