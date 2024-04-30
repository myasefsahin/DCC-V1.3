using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using System.ComponentModel;
using LicenseContext = OfficeOpenXml.LicenseContext;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DCC
{
    public class SP_DataWord
    {
        List<string> columnName = new List<string>(104) {
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
                "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ","AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
                "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ","BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT","BU", "BV", "BW", "BX", "BY", "BZ",
                "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ","CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT","CU", "CV", "CW", "CX", "CY", "CZ" };



        public void main(string ExcelDosyaYolu, string pageName,int satır,string sütun)
        {
            int harfIndex = columnName.IndexOf(sütun);
            using (var package = new ExcelPackage(new FileInfo(ExcelDosyaYolu)))
            {


                // Excel'in 13. sayfasındaki veriler

                ExcelWorksheet worksheet = package.Workbook.Worksheets[pageName];

                int rowCount = worksheet.Dimension.Rows;


                string[] cellValue = new string[rowCount];


                // Frekans değerlerinin çekimi
                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[sütun + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        ArrayFrekans.Add(cellValue[i - satır]);
                    }
                }

                // S-Parametre değerlerinin çekimi
                for (int i = satır; i < ArrayFrekans.Count + satır; i++)
                {

                    //S11 değerleri için 
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex+1] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex+2] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    ArrayS11Reel.Add(formattedEntity.measurent);
                    ArrayS11ReelUnc.Add(formattedEntity.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex+3]  + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex+4] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    ArrayS11Complex.Add(formattedEntity1.measurent);
                    ArrayS11ComplexUnc.Add(formattedEntity1.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 6] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 7] + i].Value);
                    CalculateEntity formattedEntity2 = NumberFormatter.deneme(calculateEntity);
                    ArrayS11Lin.Add(formattedEntity2.measurent);
                    ArrayS11LinUnc.Add(formattedEntity2.uncertainty);


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 8] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 9] + i].Value);
                    CalculateEntity formattedEntity3 = NumberFormatter.deneme(calculateEntity);
                    ArrayS11LinPhase.Add(formattedEntity3.measurent);
                    ArrayS11LinPhaseUnc.Add(formattedEntity3.uncertainty);


                    calculateEntity.measurent = Math.Abs(Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 11] + i].Value));
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 12] + i].Value);
                    CalculateEntity formattedEntity4 = NumberFormatter.deneme(calculateEntity);
                    ArrayS11Log.Add(formattedEntity4.measurent);
                    ArrayS11LogUnc.Add(formattedEntity4.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 13] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 14] + i].Value);
                    CalculateEntity formattedEntity5 = NumberFormatter.deneme(calculateEntity);
                    ArrayS11LogPhase.Add(formattedEntity5.measurent);
                    ArrayS11LogPhaseUnc.Add(formattedEntity5.uncertainty);


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 16] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 17] + i].Value);
                    CalculateEntity formattedEntity6 = NumberFormatter.deneme(calculateEntity);
                    ArrayS11SWR.Add(formattedEntity6.measurent);
                    ArrayS11SWRUnc.Add(formattedEntity6.uncertainty);

                    //S12 değerleri için 
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex +19] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 20] + i].Value);
                    CalculateEntity formattedEntity7 = NumberFormatter.deneme(calculateEntity);
                    ArrayS12Reel.Add(formattedEntity7.measurent);
                    ArrayS12ReelUnc.Add(formattedEntity7.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 21] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 22] + i].Value);
                    CalculateEntity formattedEntity8 = NumberFormatter.deneme(calculateEntity);
                    ArrayS12Complex.Add(formattedEntity8.measurent);
                    ArrayS12ComplexUnc.Add(formattedEntity8.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 24] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 25] + i].Value);
                    CalculateEntity formattedEntity9 = NumberFormatter.deneme(calculateEntity);
                    ArrayS12Lin.Add(formattedEntity9.measurent);
                    ArrayS12LinUnc.Add(formattedEntity9.uncertainty);


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 26] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 27] + i].Value);
                    CalculateEntity formattedEntity10 = NumberFormatter.deneme(calculateEntity);
                    ArrayS12LinPhase.Add(formattedEntity10.measurent);
                    ArrayS12LinPhaseUnc.Add(formattedEntity10.uncertainty);


                    calculateEntity.measurent = Math.Abs(Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 29] + i].Value));
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 30] + i].Value);
                    CalculateEntity formattedEntity11 = NumberFormatter.deneme(calculateEntity);
                    ArrayS12Log.Add(formattedEntity11.measurent);
                    ArrayS12LogUnc.Add(formattedEntity11.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 31] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 32] + i].Value);
                    CalculateEntity formattedEntity12 = NumberFormatter.deneme(calculateEntity);
                    ArrayS12LogPhase.Add(formattedEntity12.measurent);
                    ArrayS12LogPhaseUnc.Add(formattedEntity12.uncertainty);


                    //S21 değerleri için 


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 34] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 35] + i].Value);
                    CalculateEntity formattedEntity13 = NumberFormatter.deneme(calculateEntity);
                    ArrayS21Reel.Add(formattedEntity13.measurent);
                    ArrayS21ReelUnc.Add(formattedEntity13.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 36] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 37] + i].Value);
                    CalculateEntity formattedEntity14 = NumberFormatter.deneme(calculateEntity);
                    ArrayS21Complex.Add(formattedEntity14.measurent);
                    ArrayS21ComplexUnc.Add(formattedEntity14.uncertainty);


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 39] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 40] + i].Value);
                    CalculateEntity formattedEntity15 = NumberFormatter.deneme(calculateEntity);
                    ArrayS21Lin.Add(formattedEntity15.measurent);
                    ArrayS21LinUnc.Add(formattedEntity15.uncertainty);


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 41] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 42] + i].Value);
                    CalculateEntity formattedEntity16 = NumberFormatter.deneme(calculateEntity);
                    ArrayS21LinPhase.Add(formattedEntity16.measurent);
                    ArrayS21LinPhaseUnc.Add(formattedEntity16.uncertainty);


                    calculateEntity.measurent = Math.Abs(Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 44] + i].Value));
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 45] + i].Value);
                    CalculateEntity formattedEntity17 = NumberFormatter.deneme(calculateEntity);
                    ArrayS21Log.Add(formattedEntity17.measurent);
                    ArrayS21LogUnc.Add(formattedEntity17.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 46] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 47] + i].Value);
                    CalculateEntity formattedEntity18 = NumberFormatter.deneme(calculateEntity);
                    ArrayS21LogPhase.Add(formattedEntity18.measurent);
                    ArrayS21LogPhaseUnc.Add(formattedEntity18.uncertainty);


                    //S22 değerleri için 

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 49] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 50] + i].Value);
                    CalculateEntity formattedEntity19 = NumberFormatter.deneme(calculateEntity);
                    ArrayS22Reel.Add(formattedEntity19.measurent);
                    ArrayS22ReelUnc.Add(formattedEntity19.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 51] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 52] + i].Value);
                    CalculateEntity formattedEntity20 = NumberFormatter.deneme(calculateEntity);
                    ArrayS22Complex.Add(formattedEntity20.measurent);
                    ArrayS22ComplexUnc.Add(formattedEntity20.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 54] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 55] + i].Value);
                    CalculateEntity formattedEntity21 = NumberFormatter.deneme(calculateEntity);
                    ArrayS22Lin.Add(formattedEntity21.measurent);
                    ArrayS22LinUnc.Add(formattedEntity21.uncertainty);


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 56] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 57] + i].Value);
                    CalculateEntity formattedEntity22 = NumberFormatter.deneme(calculateEntity);
                    ArrayS22LinPhase.Add(formattedEntity22.measurent);
                    ArrayS22LinPhaseUnc.Add(formattedEntity22.uncertainty);


                    calculateEntity.measurent = Math.Abs(Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 59] + i].Value));
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 60] + i].Value);
                    CalculateEntity formattedEntity23 = NumberFormatter.deneme(calculateEntity);
                    ArrayS22Log.Add(formattedEntity23.measurent);
                    ArrayS22LogUnc.Add(formattedEntity23.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 61] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 62] + i].Value);
                    CalculateEntity formattedEntity24 = NumberFormatter.deneme(calculateEntity);
                    ArrayS22LogPhase.Add(formattedEntity24.measurent);
                    ArrayS22LogPhaseUnc.Add(formattedEntity24.uncertainty);


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 64] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 65] + i].Value);
                    CalculateEntity formattedEntity25 = NumberFormatter.deneme(calculateEntity);
                    ArrayS22SWR.Add(formattedEntity25.measurent);
                    ArrayS22SWRUnc.Add(formattedEntity25.uncertainty);

                }
                tableName1= Convert.ToString(worksheet.Cells[columnName[harfIndex ] + (satır-2)].Value);
                tableName2 = Convert.ToString(worksheet.Cells[columnName[harfIndex+6] + (satır - 2)].Value);
                tableName3 = Convert.ToString(worksheet.Cells[columnName[harfIndex+11] + (satır - 2)].Value);
                tableName4 = Convert.ToString(worksheet.Cells[columnName[harfIndex+16] + (satır - 2)].Value);
                tableName5 = Convert.ToString(worksheet.Cells[columnName[harfIndex+19] + (satır - 2)].Value);
                tableName6 = Convert.ToString(worksheet.Cells[columnName[harfIndex+24] + (satır - 2)].Value);
                tableName7 = Convert.ToString(worksheet.Cells[columnName[harfIndex+29] + (satır - 2)].Value);
                tableName8 = Convert.ToString(worksheet.Cells[columnName[harfIndex+34] + (satır - 2)].Value);
                tableName9 = Convert.ToString(worksheet.Cells[columnName[harfIndex+39] + (satır - 2)].Value);
                tableName10 = Convert.ToString(worksheet.Cells[columnName[harfIndex+44] + (satır - 2)].Value);
                tableName11 = Convert.ToString(worksheet.Cells[columnName[harfIndex+49] + (satır - 2)].Value);
                tableName12 = Convert.ToString(worksheet.Cells[columnName[harfIndex+54] + (satır - 2)].Value);
                tableName13 = Convert.ToString(worksheet.Cells[columnName[harfIndex+59] + (satır - 2)].Value);
                tableName14 = Convert.ToString(worksheet.Cells[columnName[harfIndex+64] + (satır - 2)].Value);













            }

        }
        public string tableName1;
        public string tableName2;
        public string tableName3;
        public string tableName4;
        public string tableName5;
        public string tableName6;
        public string tableName7;
        public string tableName8;
        public string tableName9;
        public string tableName10;
        public string tableName11;
        public string tableName12;
        public string tableName13;
        public string tableName14;



        public ArrayList ArrayFrekans { get; set; }

        // S11
        public ArrayList ArrayS11Reel { get; set; }
        public ArrayList ArrayS11ReelUnc { get; set; }
        public ArrayList ArrayS11Complex { get; set; }
        public ArrayList ArrayS11ComplexUnc { get; set; }
        public ArrayList ArrayS11Lin { get; set; }
        public ArrayList ArrayS11LinUnc { get; set; }
        public ArrayList ArrayS11LinPhase { get; set; }
        public ArrayList ArrayS11LinPhaseUnc { get; set; }
        public ArrayList ArrayS11Log { get; set; }
        public ArrayList ArrayS11LogUnc { get; set; }
        public ArrayList ArrayS11LogPhase { get; set; }
        public ArrayList ArrayS11LogPhaseUnc { get; set; }
        public ArrayList ArrayS11SWR { get; set; }
        public ArrayList ArrayS11SWRUnc { get; set; }

        // S12
        public ArrayList ArrayS12Reel { get; set; }
        public ArrayList ArrayS12ReelUnc { get; set; }
        public ArrayList ArrayS12Complex { get; set; }
        public ArrayList ArrayS12ComplexUnc { get; set; }
        public ArrayList ArrayS12Lin { get; set; }
        public ArrayList ArrayS12LinUnc { get; set; }
        public ArrayList ArrayS12LinPhase { get; set; }
        public ArrayList ArrayS12LinPhaseUnc { get; set; }
        public ArrayList ArrayS12Log { get; set; }
        public ArrayList ArrayS12LogUnc { get; set; }
        public ArrayList ArrayS12LogPhase { get; set; }
        public ArrayList ArrayS12LogPhaseUnc { get; set; }

        // S21
        public ArrayList ArrayS21Reel { get; set; }
        public ArrayList ArrayS21ReelUnc { get; set; }
        public ArrayList ArrayS21Complex { get; set; }
        public ArrayList ArrayS21ComplexUnc { get; set; }
        public ArrayList ArrayS21Lin { get; set; }
        public ArrayList ArrayS21LinUnc { get; set; }
        public ArrayList ArrayS21LinPhase { get; set; }
        public ArrayList ArrayS21LinPhaseUnc { get; set; }
        public ArrayList ArrayS21Log { get; set; }
        public ArrayList ArrayS21LogUnc { get; set; }
        public ArrayList ArrayS21LogPhase { get; set; }
        public ArrayList ArrayS21LogPhaseUnc { get; set; }

        // S22
        public ArrayList ArrayS22Reel { get; set; }
        public ArrayList ArrayS22ReelUnc { get; set; }
        public ArrayList ArrayS22Complex { get; set; }
        public ArrayList ArrayS22ComplexUnc { get; set; }
        public ArrayList ArrayS22Lin { get; set; }
        public ArrayList ArrayS22LinUnc { get; set; }
        public ArrayList ArrayS22LinPhase { get; set; }
        public ArrayList ArrayS22LinPhaseUnc { get; set; }
        public ArrayList ArrayS22Log { get; set; }
        public ArrayList ArrayS22LogUnc { get; set; }
        public ArrayList ArrayS22LogPhase { get; set; }
        public ArrayList ArrayS22LogPhaseUnc { get; set; }
        public ArrayList ArrayS22SWR { get; set; }
        public ArrayList ArrayS22SWRUnc { get; set; }

        // Device Informations

        public string OrderNumber { get; set; }
        public string DeviceName { get; set; }
        public string SerialNumber { get; set; }


        public SP_DataWord()
        {
            ArrayFrekans = new ArrayList();

            ArrayS11Reel = new ArrayList();
            ArrayS11ReelUnc = new ArrayList();
            ArrayS11Complex = new ArrayList();
            ArrayS11ComplexUnc = new ArrayList();
            ArrayS11Lin = new ArrayList();
            ArrayS11LinUnc = new ArrayList();
            ArrayS11LinPhase = new ArrayList();
            ArrayS11LinPhaseUnc = new ArrayList();
            ArrayS11Log = new ArrayList();
            ArrayS11LogUnc = new ArrayList();
            ArrayS11LogPhase = new ArrayList();
            ArrayS11LogPhaseUnc = new ArrayList();
            ArrayS11SWR = new ArrayList();
            ArrayS11SWRUnc = new ArrayList();


            ArrayS12Reel = new ArrayList();
            ArrayS12ReelUnc = new ArrayList();
            ArrayS12Complex = new ArrayList();
            ArrayS12ComplexUnc = new ArrayList();
            ArrayS12Lin = new ArrayList();
            ArrayS12LinUnc = new ArrayList();
            ArrayS12LinPhase = new ArrayList();
            ArrayS12LinPhaseUnc = new ArrayList();
            ArrayS12Log = new ArrayList();
            ArrayS12LogUnc = new ArrayList();
            ArrayS12LogPhase = new ArrayList();
            ArrayS12LogPhaseUnc = new ArrayList();


            ArrayS21Reel = new ArrayList();
            ArrayS21ReelUnc = new ArrayList();
            ArrayS21Complex = new ArrayList();
            ArrayS21ComplexUnc = new ArrayList();
            ArrayS21Lin = new ArrayList();
            ArrayS21LinUnc = new ArrayList();
            ArrayS21LinPhase = new ArrayList();
            ArrayS21LinPhaseUnc = new ArrayList();
            ArrayS21Log = new ArrayList();
            ArrayS21LogUnc = new ArrayList();
            ArrayS21LogPhase = new ArrayList();
            ArrayS21LogPhaseUnc = new ArrayList();


            ArrayS22Reel = new ArrayList();
            ArrayS22ReelUnc = new ArrayList();
            ArrayS22Complex = new ArrayList();
            ArrayS22ComplexUnc = new ArrayList();
            ArrayS22Lin = new ArrayList();
            ArrayS22LinUnc = new ArrayList();
            ArrayS22LinPhase = new ArrayList();
            ArrayS22LinPhaseUnc = new ArrayList();
            ArrayS22Log = new ArrayList();
            ArrayS22LogUnc = new ArrayList();
            ArrayS22LogPhase = new ArrayList();
            ArrayS22LogPhaseUnc = new ArrayList();
            ArrayS22SWR = new ArrayList();
            ArrayS22SWRUnc = new ArrayList();
        }

        public void ClearData()
        {
            ArrayFrekans.Clear();
            ArrayS11Reel.Clear();
            ArrayS11ReelUnc.Clear();
            ArrayS11Complex.Clear();
            ArrayS11ComplexUnc.Clear();
            ArrayS11Lin.Clear();
            ArrayS11LinUnc.Clear();
            ArrayS11LinPhase.Clear();
            ArrayS11LinPhaseUnc.Clear();
            ArrayS11Log.Clear();
            ArrayS11LogUnc.Clear();
            ArrayS11LogPhase.Clear();
            ArrayS11LogPhaseUnc.Clear();
            ArrayS11SWR.Clear();
            ArrayS11SWRUnc.Clear();
            ArrayS12Reel.Clear();
            ArrayS12ReelUnc.Clear();
            ArrayS12Complex.Clear();
            ArrayS12ComplexUnc.Clear();
            ArrayS12Lin.Clear();
            ArrayS12LinUnc.Clear();
            ArrayS12LinPhase.Clear();
            ArrayS12LinPhaseUnc.Clear();
            ArrayS12Log.Clear();
            ArrayS12LogUnc.Clear();
            ArrayS12LogPhase.Clear();
            ArrayS12LogPhaseUnc.Clear();
            ArrayS21Reel.Clear();
            ArrayS21ReelUnc.Clear();
            ArrayS21Complex.Clear();
            ArrayS21ComplexUnc.Clear();
            ArrayS21Lin.Clear();
            ArrayS21LinUnc.Clear();
            ArrayS21LinPhase.Clear();
            ArrayS21LinPhaseUnc.Clear();
            ArrayS21Log.Clear();
            ArrayS21LogUnc.Clear();
            ArrayS21LogPhase.Clear();
            ArrayS21LogPhaseUnc.Clear();
            ArrayS22Reel.Clear();
            ArrayS22ReelUnc.Clear();
            ArrayS22Complex.Clear();
            ArrayS22ComplexUnc.Clear();
            ArrayS22Lin.Clear();
            ArrayS22LinUnc.Clear();
            ArrayS22LinPhase.Clear();
            ArrayS22LinPhaseUnc.Clear();
            ArrayS22Log.Clear();
            ArrayS22LogUnc.Clear();
            ArrayS22LogPhase.Clear();
            ArrayS22LogPhaseUnc.Clear();
            ArrayS22SWR.Clear();
            ArrayS22SWRUnc.Clear();
        }
    }
}



