
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
        public void main(string ExcelDosyaYolu, string pageName)
        {
            using (var package = new ExcelPackage(new FileInfo(ExcelDosyaYolu)))
            {


                // Excel'in sabit olan sonuç sayfasındaki verilere göre hazırlanmıştır. 
                ExcelWorksheet worksheet = package.Workbook.Worksheets[pageName];
                int rowCount = worksheet.Dimension.Rows;
                //int satirSayisi = rowCount - 6;

                string[] cellValue = new string[rowCount];


                // Frekans değerlerinin çekimi
                for (int i = 7; i <= rowCount; i++)
                {
                    cellValue[i - 7] = Convert.ToString(worksheet.Cells["N" + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - 7]))
                    {
                        CIS_Olcum_Adım.Add(cellValue[i - 7]);
                    }
                }
                for (int i = 7; i < CIS_Olcum_Adım.Count + 7; i++)
                {

                    //S11 değerleri için 
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();



                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells["O" + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells["P" + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    CIS_ZP.Add(formattedEntity.measurent);
                    CIS_ZP_Unc.Add(formattedEntity.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells["Q" + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells["R" + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    CIS_ICOD.Add(formattedEntity1.measurent);
                    CIS_ICOD_Unc.Add(formattedEntity1.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells["S" + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells["T" + i].Value);
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
