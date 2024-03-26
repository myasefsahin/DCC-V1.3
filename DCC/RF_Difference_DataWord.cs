using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace DCC
{
     class RF_Difference_DataWord
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

                ExcelWorksheet worksheet = package.Workbook.Worksheets[pageName];
                int rowCount = worksheet.Dimension.End.Row;
                string[] cellValue = new string[rowCount];


                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[sütun + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        RFD_T1_Frekans.Add(cellValue[i - satır]);
                    }
                }


                for (int i = satır; i < RFD_T1_Frekans.Count + satır; i++)
                {

                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    RFD_T1_AltSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 2] + i].Value));
                    RFD_T1_ÜstSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 5] + i].Value));

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 1] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 6] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    RFD_T1_GostergeDegeri.Add(formattedEntity.measurent);
                    RFD_T1_Belirsizlik.Add(formattedEntity.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 3] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 6] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    RFD_T1_OlculenDeger.Add(formattedEntity1.measurent);


                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 4] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 6] + i].Value);
                    CalculateEntity formattedEntity2 = NumberFormatter.deneme(calculateEntity);
                    RFD_T1_OlculenFark.Add(formattedEntity2.measurent);



                }
              

                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 8] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        RFD_T2_Frekans.Add(cellValue[i - satır]);
                    }
                }


                for (int i = satır; i < RFD_T2_Frekans.Count + satır; i++)
                {

                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    RFD_T2_Nom_Guc_Lvl.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 9] + i].Value));
                    RFD_T2_AltSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 11] + i].Value));
                    RFD_T2_ÜstSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 13] + i].Value));

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 10] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 14] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    RFD_T2_OlculenDeger.Add(formattedEntity.measurent);
                    RFD_T2_Belirsizlik.Add(formattedEntity.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 12] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 14] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    RFD_T2_Nom_Guc_Lvl_fark.Add(formattedEntity1.measurent);



                }


                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 16] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        RFD_T3_Frekans.Add(cellValue[i - satır]);
                    }
                }


                for (int i = 7; i < RFD_T3_Frekans.Count + satır; i++)
                {

                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    RFD_T3_NominalGuc.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 17] + i].Value));
                    RFD_T3_AltSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 18] + i].Value));
                    RFD_T3_ÜstSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 20] + i].Value));

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 19] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 22] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    RFD_T3_OlculenDeger.Add(formattedEntity.measurent);
                    RFD_T3_Belirsizlik.Add(formattedEntity.uncertainty);

                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 21] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 22] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    RFD_T3_Fark.Add(formattedEntity1.measurent);


                }


                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[columnName[harfIndex + 26] + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        RFD_T4_Frekans.Add(cellValue[i - satır]);
                    }
                }


                for (int i = 7; i < RFD_T4_Frekans.Count + satır; i++)
                {
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    RFD_T4_Min_Guc_lvl.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 24] + i].Value));
                    RFD_T4_Max_Guc_lvl.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 25] + i].Value));

                    RFD_T4_AltSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 27] + i].Value));
                    RFD_T4_UstSınır.Add(Convert.ToString(worksheet.Cells[columnName[harfIndex + 29] + i].Value));



                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 28] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 30] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    RFD_T4_Fark.Add(formattedEntity1.measurent);
                    RFD_T4_Belirsizlik.Add(formattedEntity1.uncertainty);



                }
            }
        }


        public ArrayList RFD_T1_Frekans { get; set; }
        public ArrayList RFD_T1_GostergeDegeri { get; set; }
        public ArrayList RFD_T1_AltSınır { get; set; }
        public ArrayList RFD_T1_OlculenDeger { get; set; }
        public ArrayList RFD_T1_OlculenFark { get; set; }
        public ArrayList RFD_T1_ÜstSınır { get; set; }
        public ArrayList RFD_T1_Belirsizlik { get; set; }


        public ArrayList RFD_T2_Frekans { get; set; }
        public ArrayList RFD_T2_Nom_Guc_Lvl { get; set; }
        public ArrayList RFD_T2_OlculenDeger { get; set; }
        public ArrayList RFD_T2_AltSınır { get; set; }
        public ArrayList RFD_T2_Nom_Guc_Lvl_fark { get; set; }
        public ArrayList RFD_T2_ÜstSınır { get; set; }
        public ArrayList RFD_T2_Belirsizlik { get; set; }


        public ArrayList RFD_T3_Frekans { get; set; }
        public ArrayList RFD_T3_NominalGuc { get; set; }
        public ArrayList RFD_T3_AltSınır { get; set; }
        public ArrayList RFD_T3_OlculenDeger { get; set; }
        public ArrayList RFD_T3_ÜstSınır { get; set; }
        public ArrayList RFD_T3_Fark { get; set; } 
        public ArrayList RFD_T3_Belirsizlik { get; set; }



        public ArrayList RFD_T4_Min_Guc_lvl { get; set; }
        public ArrayList RFD_T4_Max_Guc_lvl { get; set; }
        public ArrayList RFD_T4_Frekans { get; set; }
        public ArrayList RFD_T4_AltSınır { get; set; }
        public ArrayList RFD_T4_Fark { get; set; }
        public ArrayList RFD_T4_UstSınır { get; set; }
        public ArrayList RFD_T4_Belirsizlik { get; set; }


        public RF_Difference_DataWord()
        {
            RFD_T1_Frekans = new ArrayList();
            RFD_T1_GostergeDegeri = new ArrayList();
            RFD_T1_AltSınır = new ArrayList();
            RFD_T1_OlculenDeger = new ArrayList();
            RFD_T1_OlculenFark = new ArrayList();
            RFD_T1_ÜstSınır = new ArrayList();
            RFD_T1_Belirsizlik = new ArrayList();

            RFD_T2_Frekans = new ArrayList();
            RFD_T2_Nom_Guc_Lvl = new ArrayList();
            RFD_T2_OlculenDeger = new ArrayList();
            RFD_T2_AltSınır = new ArrayList();
            RFD_T2_Nom_Guc_Lvl_fark = new ArrayList();
            RFD_T2_ÜstSınır = new ArrayList();
            RFD_T2_Belirsizlik = new ArrayList();

            RFD_T3_Frekans = new ArrayList();
            RFD_T3_NominalGuc = new ArrayList();
            RFD_T3_AltSınır = new ArrayList();
            RFD_T3_OlculenDeger = new ArrayList();
            RFD_T3_ÜstSınır = new ArrayList();
            RFD_T3_Fark = new ArrayList();
            RFD_T3_Belirsizlik = new ArrayList();

            RFD_T4_Min_Guc_lvl = new ArrayList();
            RFD_T4_Max_Guc_lvl = new ArrayList();
            RFD_T4_Frekans = new ArrayList();
            RFD_T4_AltSınır = new ArrayList();
            RFD_T4_Fark = new ArrayList();
            RFD_T4_UstSınır = new ArrayList();
            RFD_T4_Belirsizlik = new ArrayList();


        }

        public void ClearData()
        {
            RFD_T1_Frekans.Clear();
            RFD_T1_GostergeDegeri.Clear();
            RFD_T1_AltSınır.Clear();
            RFD_T1_OlculenDeger.Clear();
            RFD_T1_OlculenFark.Clear();
            RFD_T1_ÜstSınır.Clear();
            RFD_T1_Belirsizlik.Clear();

            RFD_T2_Frekans.Clear();
            RFD_T2_Nom_Guc_Lvl.Clear();
            RFD_T2_OlculenDeger.Clear();
            RFD_T2_AltSınır.Clear();
            RFD_T2_Nom_Guc_Lvl_fark.Clear();
            RFD_T2_ÜstSınır.Clear();
            RFD_T2_Belirsizlik.Clear();

            RFD_T3_Frekans.Clear();
            RFD_T3_NominalGuc.Clear();
            RFD_T3_AltSınır.Clear();
            RFD_T3_OlculenDeger.Clear();
            RFD_T3_ÜstSınır.Clear();
            RFD_T3_Fark.Clear();
            RFD_T3_Belirsizlik.Clear(); 

            RFD_T4_Min_Guc_lvl.Clear();
            RFD_T4_Max_Guc_lvl.Clear();
            RFD_T4_Frekans.Clear();
            RFD_T4_AltSınır.Clear();
            RFD_T4_Fark.Clear();
            RFD_T4_UstSınır.Clear();
            RFD_T4_Belirsizlik.Clear();
        }
    }
}
