using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
     class Noise_DataWord
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
                        NS_ArrayFrekans.Add(cellValue[i - satır]);
                    }
                }


                // S-Parametre değerlerinin çekimi
                for (int i = satır; i < NS_ArrayFrekans.Count + satır; i++)
                {


                    //S11 değerleri için 
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    
                  
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 1] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 2] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    NS_ArrayENR.Add(formattedEntity.measurent);
                    NS_ArrayENRUnc.Add(formattedEntity.uncertainty);


                    

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 4] + i].Value);
                    NS_ArrayRC_ustlimit.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 5] + i].Value));
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 6] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    NS_ArrayRC.Add(formattedEntity1.measurent);
                    NS_ArrayRCUnc.Add(formattedEntity1.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 7] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 8] + i].Value);
                    CalculateEntity formattedEntity7 = NumberFormatter.deneme(calculateEntity);
                    NS_ArrayRC_Phase.Add(formattedEntity7.measurent);
                    NS_ArrayRC_PhaseUnc.Add(formattedEntity7.uncertainty);
                    NS_ArrayControl_DC_ON.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 9] + i].Value));




                    

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 11] + i].Value);
                    NS_ArrayRC_ustlimit_DC_OFF.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 12] + i].Value));
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 13] + i].Value);
                    CalculateEntity formattedEntity8 = NumberFormatter.deneme(calculateEntity);
                    NS_ArrayRC_DC_OFF.Add(formattedEntity8.measurent);
                    NS_ArrayRCUnc_DC_OFF.Add(formattedEntity8.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 14] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 15] + i].Value);
                    CalculateEntity formattedEntity13 = NumberFormatter.deneme(calculateEntity);
                    NS_ArrayRC_Phase_DC_OFF.Add(formattedEntity13.measurent);
                    NS_ArrayRC_PhaseUnc_DC_OFF.Add(formattedEntity13.uncertainty);

                    NS_ArrayControl_DC_OFF.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 16] + i].Value));



                }
                tableName1= Convert.ToString(worksheet.Cells[columnName[harfIndex] + (satır-2)].Value);
                tableName2 = Convert.ToString(worksheet.Cells[columnName[harfIndex+4] + (satır - 2)].Value);
                tableName3 = Convert.ToString(worksheet.Cells[columnName[harfIndex+11] + (satır - 2)].Value);

            }

        }
        public string tableName1;
        public string tableName2;
        public string tableName3;
        public ArrayList NS_ArrayFrekans { get; set; }

        public ArrayList NS_ArrayENR { get; set; }
        public ArrayList NS_ArrayENRUnc { get; set; }

        public ArrayList NS_ArrayRC { get; set; }
        public ArrayList NS_ArrayRC_ustlimit { get; set; }
        public ArrayList NS_ArrayRCUnc { get; set; }
        public ArrayList NS_ArrayRC_Phase { get; set; }
        public ArrayList NS_ArrayRC_PhaseUnc { get; set; }
        public ArrayList NS_ArrayControl_DC_ON { get; set; }



     
        public ArrayList NS_ArrayRC_DC_OFF { get; set; }
        public ArrayList NS_ArrayRC_ustlimit_DC_OFF { get; set; }
        public ArrayList NS_ArrayRCUnc_DC_OFF { get; set; }
        public ArrayList NS_ArrayRC_Phase_DC_OFF { get; set; }
        public ArrayList NS_ArrayRC_PhaseUnc_DC_OFF { get; set; }
        public ArrayList NS_ArrayControl_DC_OFF { get; set; }





        public Noise_DataWord()
        {
            NS_ArrayFrekans = new ArrayList();

            NS_ArrayENR = new ArrayList();
            NS_ArrayENRUnc = new ArrayList();

            NS_ArrayRC = new ArrayList();
            NS_ArrayRC_ustlimit = new ArrayList();
            NS_ArrayRCUnc = new ArrayList();
            NS_ArrayRC_Phase = new ArrayList();
            NS_ArrayRC_PhaseUnc = new ArrayList();
            NS_ArrayControl_DC_ON = new ArrayList();

            NS_ArrayRC_DC_OFF = new ArrayList();
            NS_ArrayRC_ustlimit_DC_OFF = new ArrayList();
            NS_ArrayRCUnc_DC_OFF = new ArrayList();
            NS_ArrayRC_Phase_DC_OFF = new ArrayList();
            NS_ArrayRC_PhaseUnc_DC_OFF = new ArrayList();
            NS_ArrayControl_DC_OFF = new ArrayList();


        }

        public void ClearData()
        {
            NS_ArrayFrekans.Clear();
            NS_ArrayENR.Clear();
            NS_ArrayENRUnc.Clear();
            NS_ArrayRC.Clear();
            NS_ArrayRC_ustlimit.Clear();
            NS_ArrayRCUnc.Clear();
            NS_ArrayRC_Phase.Clear();
            NS_ArrayRC_PhaseUnc.Clear();
            NS_ArrayControl_DC_ON.Clear();
            NS_ArrayRC_DC_OFF.Clear();
            NS_ArrayRC_ustlimit_DC_OFF.Clear();
            NS_ArrayRCUnc_DC_OFF.Clear();
            NS_ArrayRC_Phase_DC_OFF.Clear();
            NS_ArrayRC_PhaseUnc_DC_OFF.Clear();
            NS_ArrayControl_DC_OFF.Clear();
           
        }
    }
}
