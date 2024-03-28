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
    public class Absolute_RF_Power_DataWord

    {
        List<string> columnName = new List<string>(104) {
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
                "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ","AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
                "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ","BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT","BU", "BV", "BW", "BX", "BY", "BZ",
                "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ","CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT","CU", "CV", "CW", "CX", "CY", "CZ" };
        CertificateForm certificate;


        public void main(string ExcelDosyaYolu, string pageName, int satır, string sütun)
        {



            int harfIndex = columnName.IndexOf(sütun);
            using (var package = new ExcelPackage(new FileInfo(ExcelDosyaYolu)))
            {


                // Excel'in 13. sayfasındaki veriler

                ExcelWorksheet worksheet = package.Workbook.Worksheets[pageName];

                int rowCount = worksheet.Dimension.End.Row;


                string[] cellValue = new string[rowCount];


                // Frekans değerlerinin çekimi
                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[sütun + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        ARFP_T1_Frekans.Add(cellValue[i - satır]);
                    }
                }

                // S-Parametre değerlerinin çekimi
                for (int i = satır; i < ARFP_T1_Frekans.Count + satır; i++)
                {

                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    ARFP_T1_Cıkıs_Gücü.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 1] + i].Value));
                    
                    ARFP_T1_AltSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 3] + i].Value));
                    ARFP_T1_ÜstSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 5] + i].Value));

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 4] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 6] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);

                    ARFP_T1_Sapma.Add(formattedEntity.measurent);
                    ARFP_T1_Belirsizlik.Add(formattedEntity.uncertainty);




                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 2] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 6] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);

                    ARFP_T1_Olculen_Güc.Add(formattedEntity1.measurent);
                    
                }

                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 8] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        ARFP_T2_Frekans.Add(cellValue[i - satır]);
                    }
                }


                for (int i = satır; i < ARFP_T2_Frekans.Count + satır; i++)
                {

                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    ARFP_T2_Cıkıs_Gücü.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 9] + i].Value));
                    
                    ARFP_T2_AltSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 11] + i].Value));
                    ARFP_T2_ÜstSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 13] + i].Value));
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 12] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 14] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    ARFP_T2_Fark.Add(formattedEntity.measurent);
                    ARFP_T2_Belirsizlik.Add(formattedEntity.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 10] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 14] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    ARFP_T2_OlculenDeger.Add(formattedEntity1.measurent);
                   

                }


                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 16] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        ARFP_T3_Frekans.Add(cellValue[i - satır]);
                    }
                }


                for (int i = satır; i < ARFP_T3_Frekans.Count + satır; i++)
                {

                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    ARFP_T3_Cıkıs_Gücü.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 17] + i].Value));
                   
                    ARFP_T3_AltSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 19] + i].Value));
                    ARFP_T3_ÜstSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 21] + i].Value));
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 20] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 22] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    ARFP_T3_Zayıflatma.Add(formattedEntity.measurent);
                    ARFP_T3_Belirsizlik.Add(formattedEntity.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 18] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 22] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    ARFP_T3_OlculenZayıflatma.Add(formattedEntity1.measurent);

                }


                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 24] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        ARFP_T4_T5_T6_frekans.Add(cellValue[i - satır]);
                    }
                }


                for (int i = satır; i < ARFP_T4_T5_T6_frekans.Count + satır; i++)
                {
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    ARFP_T4_SWR_Seviye.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 25] + i].Value));
                    ARFP_T4_SWR_MaksimumDeger.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 27] + i].Value));
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 26] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 28] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    ARFP_T4_SWR_OlculenDeger.Add(formattedEntity.measurent);
                    ARFP_T4_SWR_Belirsizlik.Add(formattedEntity.uncertainty);


                    ARFP_T5_SWR_Seviye.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 31] + i].Value));
                    ARFP_T5_SWR_MaksimumDeger.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 33] + i].Value));
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 32] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 34] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    ARFP_T5_SWR_OlculenDeger.Add(formattedEntity1.measurent);
                    ARFP_T5_SWR_Belirsizlik.Add(formattedEntity1.uncertainty);

                    ARFP_T6_SWR_Seviye.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 37] + i].Value));
                    ARFP_T6_SWR_MaksimumDeger.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 39] + i].Value));
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 38] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 40] + i].Value);
                    CalculateEntity formattedEntity2 = NumberFormatter.deneme(calculateEntity);
                    ARFP_T6_SWR_OlculenDeger.Add(formattedEntity2.measurent);
                    ARFP_T6_SWR_Belirsizlik.Add(formattedEntity2.uncertainty);

                }


                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 42] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        ARFP_T7_Frekans.Add(cellValue[i - satır]);
                    }
                }


                for (int i = satır; i < ARFP_T7_Frekans.Count + satır; i++)
                {
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    ARFP_T7_Cıkıs_Gücü.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 43] + i].Value));
                    
                    ARFP_T7_AltSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 45] + i].Value));
                    ARFP_T7_ÜstSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 47] + i].Value));

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 46] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 48] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    ARFP_T7_Sapma.Add(formattedEntity.measurent);
                    ARFP_T7_Belirsizlik.Add(formattedEntity.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 44] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 48] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    ARFP_T7_OlculenGuc.Add(formattedEntity1.measurent);
                   






                }

                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 50] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        ARFP_T8_Frekans.Add(cellValue[i - satır]);
                    }
                }


                for (int i = satır; i < ARFP_T8_Frekans.Count + satır; i++)
                {
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    ARFP_T8_Cıkıs_Gücü.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 51] + i].Value));
                    
                    ARFP_T8_AltSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 53] + i].Value));
                    ARFP_T8_ÜstSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 55] + i].Value));

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 54] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 56] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    ARFP_T8_Fark.Add(formattedEntity.measurent);
                    ARFP_T8_Belirsizlik.Add(formattedEntity.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 52] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 56] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    ARFP_T8_OlculenDeger.Add(formattedEntity1.measurent);

                }

                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 58] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        ARFP_T9_T10_T11_frekans.Add(cellValue[i - satır]);
                    }
                }


                for (int i = satır; i < ARFP_T9_T10_T11_frekans.Count + satır; i++)
                {
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    ARFP_T9_SWR_Seviye.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 59] + i].Value));
                    ARFP_T9_SWR_MaksimumDeger.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 61] + i].Value));
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 60] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 62] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    ARFP_T9_SWR_OlculenDeger.Add(formattedEntity.measurent);
                    ARFP_T9_SWR_Belirsizlik.Add(formattedEntity.uncertainty);

                    ARFP_T10_SWR_Seviye.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 65] + i].Value));
                    ARFP_T10_SWR_MaksimumDeger.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 67] + i].Value));
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 66] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 68] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    ARFP_T10_SWR_OlculenDeger.Add(formattedEntity1.measurent);
                    ARFP_T10_SWR_Belirsizlik.Add(formattedEntity1.uncertainty);

                    ARFP_T11_SWR_Seviye.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 71] + i].Value));
                    ARFP_T11_SWR_MaksimumDeger.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 73] + i].Value));
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 72] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 74] + i].Value);
                    CalculateEntity formattedEntity2 = NumberFormatter.deneme(calculateEntity);
                    ARFP_T11_SWR_OlculenDeger.Add(formattedEntity2.measurent);
                    ARFP_T11_SWR_Belirsizlik.Add(formattedEntity2.uncertainty);


                }





            }
        }
        public ArrayList ARFP_T1_Frekans { get; set; }

        // S11
        public ArrayList ARFP_T1_Cıkıs_Gücü { get; set; }
        public ArrayList ARFP_T1_Olculen_Güc { get; set; }


        public ArrayList ARFP_T1_AltSınır { get; set; }
        public ArrayList ARFP_T1_Sapma { get; set; }
        public ArrayList ARFP_T1_ÜstSınır { get; set; }
        public ArrayList ARFP_T1_Belirsizlik { get; set; }


        public ArrayList ARFP_T2_Frekans { get; set; }
        public ArrayList ARFP_T2_Cıkıs_Gücü { get; set; }
        public ArrayList ARFP_T2_OlculenDeger { get; set; }
        public ArrayList ARFP_T2_AltSınır { get; set; }
        public ArrayList ARFP_T2_Fark { get; set; }
        public ArrayList ARFP_T2_ÜstSınır { get; set; }
        public ArrayList ARFP_T2_Belirsizlik { get; set; }


        public ArrayList ARFP_T3_Frekans { get; set; }
        public ArrayList ARFP_T3_Cıkıs_Gücü { get; set; }
        public ArrayList ARFP_T3_OlculenZayıflatma { get; set; }
        public ArrayList ARFP_T3_AltSınır { get; set; }
        public ArrayList ARFP_T3_Zayıflatma { get; set; }
        public ArrayList ARFP_T3_ÜstSınır { get; set; }
        public ArrayList ARFP_T3_Belirsizlik { get; set; }



        public ArrayList ARFP_T4_T5_T6_frekans { get; set; }
        public ArrayList ARFP_T4_SWR_Seviye { get; set; }
        public ArrayList ARFP_T4_SWR_OlculenDeger { get; set; }
        public ArrayList ARFP_T4_SWR_MaksimumDeger { get; set; }
        public ArrayList ARFP_T4_SWR_Belirsizlik { get; set; }


        public ArrayList ARFP_T5_SWR_Seviye { get; set; }
        public ArrayList ARFP_T5_SWR_OlculenDeger { get; set; }
        public ArrayList ARFP_T5_SWR_MaksimumDeger { get; set; }
        public ArrayList ARFP_T5_SWR_Belirsizlik { get; set; }


        public ArrayList ARFP_T6_SWR_Seviye { get; set; }
        public ArrayList ARFP_T6_SWR_OlculenDeger { get; set; }
        public ArrayList ARFP_T6_SWR_MaksimumDeger { get; set; }
        public ArrayList ARFP_T6_SWR_Belirsizlik { get; set; }

        public ArrayList ARFP_T7_Frekans { get; set; }
        public ArrayList ARFP_T7_Cıkıs_Gücü { get; set; }
        public ArrayList ARFP_T7_OlculenGuc { get; set; }
        public ArrayList ARFP_T7_AltSınır { get; set; }
        public ArrayList ARFP_T7_Sapma { get; set; }
        public ArrayList ARFP_T7_ÜstSınır { get; set; }
        public ArrayList ARFP_T7_Belirsizlik { get; set; }

        public ArrayList ARFP_T8_Frekans { get; set; }
        public ArrayList ARFP_T8_Cıkıs_Gücü { get; set; }
        public ArrayList ARFP_T8_OlculenDeger { get; set; }
        public ArrayList ARFP_T8_AltSınır { get; set; }
        public ArrayList ARFP_T8_Fark { get; set; }
        public ArrayList ARFP_T8_ÜstSınır { get; set; }
        public ArrayList ARFP_T8_Belirsizlik { get; set; }

        public ArrayList ARFP_T9_T10_T11_frekans { get; set; }
        public ArrayList ARFP_T9_SWR_Seviye { get; set; }
        public ArrayList ARFP_T9_SWR_OlculenDeger { get; set; }
        public ArrayList ARFP_T9_SWR_MaksimumDeger { get; set; }
        public ArrayList ARFP_T9_SWR_Belirsizlik { get; set; }

        public ArrayList ARFP_T10_SWR_Seviye { get; set; }
        public ArrayList ARFP_T10_SWR_OlculenDeger { get; set; }
        public ArrayList ARFP_T10_SWR_MaksimumDeger { get; set; }
        public ArrayList ARFP_T10_SWR_Belirsizlik { get; set; }


        public ArrayList ARFP_T11_SWR_Seviye { get; set; }
        public ArrayList ARFP_T11_SWR_OlculenDeger { get; set; }
        public ArrayList ARFP_T11_SWR_MaksimumDeger { get; set; }
        public ArrayList ARFP_T11_SWR_Belirsizlik { get; set; }



        public Absolute_RF_Power_DataWord()
        {
            ARFP_T1_Frekans = new ArrayList();
            ARFP_T1_Cıkıs_Gücü = new ArrayList();
            ARFP_T1_Olculen_Güc = new ArrayList();
            ARFP_T1_AltSınır = new ArrayList();
            ARFP_T1_Sapma = new ArrayList();
            ARFP_T1_ÜstSınır = new ArrayList();
            ARFP_T1_Belirsizlik = new ArrayList();


            ARFP_T2_Frekans = new ArrayList();
            ARFP_T2_Cıkıs_Gücü = new ArrayList();
            ARFP_T2_OlculenDeger = new ArrayList();
            ARFP_T2_AltSınır = new ArrayList();
            ARFP_T2_Fark = new ArrayList();
            ARFP_T2_ÜstSınır = new ArrayList();
            ARFP_T2_Belirsizlik = new ArrayList();


            ARFP_T3_Frekans = new ArrayList();
            ARFP_T3_Cıkıs_Gücü = new ArrayList();
            ARFP_T3_OlculenZayıflatma = new ArrayList();
            ARFP_T3_AltSınır = new ArrayList();
            ARFP_T3_Zayıflatma = new ArrayList();
            ARFP_T3_ÜstSınır = new ArrayList();
            ARFP_T3_Belirsizlik = new ArrayList();


            ARFP_T4_T5_T6_frekans = new ArrayList();
            ARFP_T4_SWR_Seviye = new ArrayList();
            ARFP_T4_SWR_OlculenDeger = new ArrayList();
            ARFP_T4_SWR_MaksimumDeger = new ArrayList();
            ARFP_T4_SWR_Belirsizlik = new ArrayList();


            ARFP_T5_SWR_Seviye = new ArrayList();
            ARFP_T5_SWR_OlculenDeger = new ArrayList();
            ARFP_T5_SWR_MaksimumDeger = new ArrayList();
            ARFP_T5_SWR_Belirsizlik = new ArrayList();


            ARFP_T6_SWR_Seviye = new ArrayList();
            ARFP_T6_SWR_OlculenDeger = new ArrayList();
            ARFP_T6_SWR_MaksimumDeger = new ArrayList();
            ARFP_T6_SWR_Belirsizlik = new ArrayList();


            ARFP_T7_Frekans = new ArrayList();
            ARFP_T7_Cıkıs_Gücü = new ArrayList();
            ARFP_T7_OlculenGuc = new ArrayList();
            ARFP_T7_AltSınır = new ArrayList();
            ARFP_T7_Sapma = new ArrayList();
            ARFP_T7_ÜstSınır = new ArrayList();
            ARFP_T7_Belirsizlik = new ArrayList();


            ARFP_T8_Frekans = new ArrayList();
            ARFP_T8_Cıkıs_Gücü = new ArrayList();
            ARFP_T8_OlculenDeger = new ArrayList();
            ARFP_T8_AltSınır = new ArrayList();
            ARFP_T8_Fark = new ArrayList();
            ARFP_T8_ÜstSınır = new ArrayList();
            ARFP_T8_Belirsizlik = new ArrayList();


            ARFP_T9_T10_T11_frekans = new ArrayList();
            ARFP_T9_SWR_Seviye = new ArrayList();
            ARFP_T9_SWR_OlculenDeger = new ArrayList();
            ARFP_T9_SWR_MaksimumDeger = new ArrayList();
            ARFP_T9_SWR_Belirsizlik = new ArrayList();


            ARFP_T10_SWR_Seviye = new ArrayList();
            ARFP_T10_SWR_OlculenDeger = new ArrayList();
            ARFP_T10_SWR_MaksimumDeger = new ArrayList();
            ARFP_T10_SWR_Belirsizlik = new ArrayList();

            ARFP_T11_SWR_Seviye = new ArrayList();
            ARFP_T11_SWR_OlculenDeger = new ArrayList();
            ARFP_T11_SWR_MaksimumDeger = new ArrayList();
            ARFP_T11_SWR_Belirsizlik = new ArrayList();
        }

        public void ClearData()
        {
            ARFP_T1_Frekans.Clear();
            ARFP_T1_Cıkıs_Gücü.Clear();
            ARFP_T1_Olculen_Güc.Clear();
            ARFP_T1_AltSınır.Clear();
            ARFP_T1_Sapma.Clear();
            ARFP_T1_ÜstSınır.Clear();
            ARFP_T1_Belirsizlik.Clear();


            ARFP_T2_Frekans.Clear();
            ARFP_T2_Cıkıs_Gücü.Clear();
            ARFP_T2_OlculenDeger.Clear();
            ARFP_T2_AltSınır.Clear();
            ARFP_T2_Fark.Clear();
            ARFP_T2_ÜstSınır.Clear();
            ARFP_T2_Belirsizlik.Clear();


            ARFP_T3_Frekans.Clear();
            ARFP_T3_Cıkıs_Gücü.Clear();
            ARFP_T3_OlculenZayıflatma.Clear();
            ARFP_T3_AltSınır.Clear();
            ARFP_T3_Zayıflatma.Clear();
            ARFP_T3_ÜstSınır.Clear();
            ARFP_T3_Belirsizlik.Clear();



            ARFP_T4_T5_T6_frekans.Clear();
            ARFP_T4_SWR_Seviye.Clear();
            ARFP_T4_SWR_OlculenDeger.Clear();
            ARFP_T4_SWR_MaksimumDeger.Clear();
            ARFP_T4_SWR_Belirsizlik.Clear();


            ARFP_T5_SWR_Seviye.Clear();
            ARFP_T5_SWR_OlculenDeger.Clear();
            ARFP_T5_SWR_MaksimumDeger.Clear();
            ARFP_T5_SWR_Belirsizlik.Clear();


            ARFP_T6_SWR_Seviye.Clear();
            ARFP_T6_SWR_OlculenDeger.Clear();
            ARFP_T6_SWR_MaksimumDeger.Clear();
            ARFP_T6_SWR_Belirsizlik.Clear();


            ARFP_T7_Frekans.Clear();
            ARFP_T7_Cıkıs_Gücü.Clear();
            ARFP_T7_OlculenGuc.Clear();
            ARFP_T7_AltSınır.Clear();
            ARFP_T7_Sapma.Clear();
            ARFP_T7_ÜstSınır.Clear();
            ARFP_T7_Belirsizlik.Clear();


            ARFP_T8_Frekans.Clear();
            ARFP_T8_Cıkıs_Gücü.Clear();
            ARFP_T8_OlculenDeger.Clear();
            ARFP_T8_AltSınır.Clear();
            ARFP_T8_Fark.Clear();
            ARFP_T8_ÜstSınır.Clear();
            ARFP_T8_Belirsizlik.Clear();


            ARFP_T9_T10_T11_frekans.Clear();
            ARFP_T9_SWR_Seviye.Clear();
            ARFP_T9_SWR_OlculenDeger.Clear();
            ARFP_T9_SWR_MaksimumDeger.Clear();
            ARFP_T9_SWR_Belirsizlik.Clear();


            ARFP_T10_SWR_Seviye.Clear();
            ARFP_T10_SWR_OlculenDeger.Clear();
            ARFP_T10_SWR_MaksimumDeger.Clear();
            ARFP_T10_SWR_Belirsizlik.Clear();


            ARFP_T11_SWR_Seviye.Clear();
            ARFP_T11_SWR_OlculenDeger.Clear();
            ARFP_T11_SWR_MaksimumDeger.Clear();
            ARFP_T11_SWR_Belirsizlik.Clear();
        }

    }




}
