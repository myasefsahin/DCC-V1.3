using DocumentFormat.OpenXml.Drawing.Charts;
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
    class XML_Arrays
    {
        List<string> columnName = new List<string>(104) {
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
                "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ","AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
                "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ","BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT","BU", "BV", "BW", "BX", "BY", "BZ",
                "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ","CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT","CU", "CV", "CW", "CX", "CY", "CZ" };

        #region S-PARAMETER
        public void SP_Data_Xml(string ExcelDosyaYolu, string pageName,int satır,string sütun)
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
                        XmlArrayFrekans.Add(cellValue[i - satır]);
                    }
                }

                // S-Parametre değerlerinin çekimi
                for (int i = satır; i < XmlArrayFrekans.Count + satır; i++)
                {

                    XmlArrayS11Reel.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 1] + i].Value));
                    XmlArrayS11ReelUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 2] + i].Value));
                    XmlArrayS11Complex.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 3] + i].Value));
                    XmlArrayS11ComplexUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 4] + i].Value));
                    XmlArrayS11Lin.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 5] + i].Value));
                    XmlArrayS11LinUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 6] + i].Value));
                    XmlArrayS11LinPhase.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 7] + i].Value));
                    XmlArrayS11LinPhaseUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 8] + i].Value));
                    XmlArrayS11Log.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 9] + i].Value));
                    XmlArrayS11LogUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 10] + i].Value));
                    XmlArrayS11LogPhase.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 11] + i].Value));
                    XmlArrayS11LogPhaseUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 12] + i].Value));
                    XmlArrayS11SWR.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 13] + i].Value));
                    XmlArrayS11SWRUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 14] + i].Value));

                    XmlArrayS12Reel.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 16] + i].Value));
                    XmlArrayS12ReelUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 17] + i].Value));
                    XmlArrayS12Complex.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 18] + i].Value));
                    XmlArrayS12ComplexUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 19] + i].Value));
                    XmlArrayS12Lin.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 20] + i].Value));
                    XmlArrayS12LinUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 21] + i].Value));
                    XmlArrayS12LinPhase.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 22] + i].Value));
                    XmlArrayS12LinPhaseUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 23] + i].Value));
                    XmlArrayS12Log.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 24] + i].Value));
                    XmlArrayS12LogUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 25] + i].Value));
                    XmlArrayS12LogPhase.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 26] + i].Value));
                    XmlArrayS12LogPhaseUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 27] + i].Value));


                    XmlArrayS21Reel.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 29] + i].Value));
                    XmlArrayS21ReelUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 30] + i].Value));
                    XmlArrayS21Complex.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 31] + i].Value));
                    XmlArrayS21ComplexUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 32] + i].Value));
                    XmlArrayS21Lin.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 33] + i].Value));
                    XmlArrayS21LinUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 34] + i].Value));
                    XmlArrayS21LinPhase.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 35] + i].Value));
                    XmlArrayS21LinPhaseUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 36] + i].Value));
                    XmlArrayS21Log.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 37] + i].Value));
                    XmlArrayS21LogUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 38] + i].Value));
                    XmlArrayS21LogPhase.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 39] + i].Value));
                    XmlArrayS21LogPhaseUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 40] + i].Value));

                    XmlArrayS22Reel.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 42] + i].Value));
                    XmlArrayS22ReelUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 43] + i].Value));
                    XmlArrayS22Complex.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 44] + i].Value));
                    XmlArrayS22ComplexUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 45] + i].Value));
                    XmlArrayS22Lin.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 46] + i].Value));
                    XmlArrayS22LinUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 47] + i].Value));
                    XmlArrayS22LinPhase.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 48] + i].Value));
                    XmlArrayS22LinPhaseUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 49] + i].Value));
                    XmlArrayS22Log.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 50] + i].Value));
                    XmlArrayS22LogUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 51] + i].Value));
                    XmlArrayS22LogPhase.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 52] + i].Value));
                    XmlArrayS22LogPhaseUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 53] + i].Value));
                    XmlArrayS22SWR.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 54] + i].Value));
                    XmlArrayS22SWRUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 55] + i].Value));
                }
            }
        }
        public ArrayList XmlArrayFrekans { get; set; }

        // S11
        public ArrayList XmlArrayS11Reel { get; set; }
        public ArrayList XmlArrayS11ReelUnc { get; set; }
        public ArrayList XmlArrayS11Complex { get; set; }
        public ArrayList XmlArrayS11ComplexUnc { get; set; }
        public ArrayList XmlArrayS11Lin { get; set; }
        public ArrayList XmlArrayS11LinUnc { get; set; }
        public ArrayList XmlArrayS11LinPhase { get; set; }
        public ArrayList XmlArrayS11LinPhaseUnc { get; set; }
        public ArrayList XmlArrayS11Log { get; set; }
        public ArrayList XmlArrayS11LogUnc { get; set; }
        public ArrayList XmlArrayS11LogPhase { get; set; }
        public ArrayList XmlArrayS11LogPhaseUnc { get; set; }
        public ArrayList XmlArrayS11SWR { get; set; }
        public ArrayList XmlArrayS11SWRUnc { get; set; }

        // S12
        public ArrayList XmlArrayS12Reel { get; set; }
        public ArrayList XmlArrayS12ReelUnc { get; set; }
        public ArrayList XmlArrayS12Complex { get; set; }
        public ArrayList XmlArrayS12ComplexUnc { get; set; }
        public ArrayList XmlArrayS12Lin { get; set; }
        public ArrayList XmlArrayS12LinUnc { get; set; }
        public ArrayList XmlArrayS12LinPhase { get; set; }
        public ArrayList XmlArrayS12LinPhaseUnc { get; set; }
        public ArrayList XmlArrayS12Log { get; set; }
        public ArrayList XmlArrayS12LogUnc { get; set; }
        public ArrayList XmlArrayS12LogPhase { get; set; }
        public ArrayList XmlArrayS12LogPhaseUnc { get; set; }

        // S21
        public ArrayList XmlArrayS21Reel { get; set; }
        public ArrayList XmlArrayS21ReelUnc { get; set; }
        public ArrayList XmlArrayS21Complex { get; set; }
        public ArrayList XmlArrayS21ComplexUnc { get; set; }
        public ArrayList XmlArrayS21Lin { get; set; }
        public ArrayList XmlArrayS21LinUnc { get; set; }
        public ArrayList XmlArrayS21LinPhase { get; set; }
        public ArrayList XmlArrayS21LinPhaseUnc { get; set; }
        public ArrayList XmlArrayS21Log { get; set; }
        public ArrayList XmlArrayS21LogUnc { get; set; }
        public ArrayList XmlArrayS21LogPhase { get; set; }
        public ArrayList XmlArrayS21LogPhaseUnc { get; set; }

        // S22
        public ArrayList XmlArrayS22Reel { get; set; }
        public ArrayList XmlArrayS22ReelUnc { get; set; }
        public ArrayList XmlArrayS22Complex { get; set; }
        public ArrayList XmlArrayS22ComplexUnc { get; set; }
        public ArrayList XmlArrayS22Lin { get; set; }
        public ArrayList XmlArrayS22LinUnc { get; set; }
        public ArrayList XmlArrayS22LinPhase { get; set; }
        public ArrayList XmlArrayS22LinPhaseUnc { get; set; }
        public ArrayList XmlArrayS22Log { get; set; }
        public ArrayList XmlArrayS22LogUnc { get; set; }
        public ArrayList XmlArrayS22LogPhase { get; set; }
        public ArrayList XmlArrayS22LogPhaseUnc { get; set; }
        public ArrayList XmlArrayS22SWR { get; set; }
        public ArrayList XmlArrayS22SWRUnc { get; set; }

        // Device Informations

        public string XmlOrderNumber { get; set; }
        public string XmlDeviceName { get; set; }
        public string XmlSerialNumber { get; set; }


        public XML_Arrays()
        {
            XmlArrayFrekans = new ArrayList();



            XmlArrayS11Reel = new ArrayList();
            XmlArrayS11ReelUnc = new ArrayList();
            XmlArrayS11Complex = new ArrayList();
            XmlArrayS11ComplexUnc = new ArrayList();
            XmlArrayS11Lin = new ArrayList();
            XmlArrayS11LinUnc = new ArrayList();
            XmlArrayS11LinPhase = new ArrayList();
            XmlArrayS11LinPhaseUnc = new ArrayList();
            XmlArrayS11Log = new ArrayList();
            XmlArrayS11LogUnc = new ArrayList();
            XmlArrayS11LogPhase = new ArrayList();
            XmlArrayS11LogPhaseUnc = new ArrayList();
            XmlArrayS11SWR = new ArrayList();
            XmlArrayS11SWRUnc = new ArrayList();


            XmlArrayS12Reel = new ArrayList();
            XmlArrayS12ReelUnc = new ArrayList();
            XmlArrayS12Complex = new ArrayList();
            XmlArrayS12ComplexUnc = new ArrayList();
            XmlArrayS12Lin = new ArrayList();
            XmlArrayS12LinUnc = new ArrayList();
            XmlArrayS12LinPhase = new ArrayList();
            XmlArrayS12LinPhaseUnc = new ArrayList();
            XmlArrayS12Log = new ArrayList();
            XmlArrayS12LogUnc = new ArrayList();
            XmlArrayS12LogPhase = new ArrayList();
            XmlArrayS12LogPhaseUnc = new ArrayList();


            XmlArrayS21Reel = new ArrayList();
            XmlArrayS21ReelUnc = new ArrayList();
            XmlArrayS21Complex = new ArrayList();
            XmlArrayS21ComplexUnc = new ArrayList();
            XmlArrayS21Lin = new ArrayList();
            XmlArrayS21LinUnc = new ArrayList();
            XmlArrayS21LinPhase = new ArrayList();
            XmlArrayS21LinPhaseUnc = new ArrayList();
            XmlArrayS21Log = new ArrayList();
            XmlArrayS21LogUnc = new ArrayList();
            XmlArrayS21LogPhase = new ArrayList();
            XmlArrayS21LogPhaseUnc = new ArrayList();


            XmlArrayS22Reel = new ArrayList();
            XmlArrayS22ReelUnc = new ArrayList();
            XmlArrayS22Complex = new ArrayList();
            XmlArrayS22ComplexUnc = new ArrayList();
            XmlArrayS22Lin = new ArrayList();
            XmlArrayS22LinUnc = new ArrayList();
            XmlArrayS22LinPhase = new ArrayList();
            XmlArrayS22LinPhaseUnc = new ArrayList();
            XmlArrayS22Log = new ArrayList();
            XmlArrayS22LogUnc = new ArrayList();
            XmlArrayS22LogPhase = new ArrayList();
            XmlArrayS22LogPhaseUnc = new ArrayList();
            XmlArrayS22SWR = new ArrayList();
            XmlArrayS22SWRUnc = new ArrayList();
            XML_EE_ArrayFrekans = new ArrayList();
            XML_EE_ArrayEE = new ArrayList();
            XML_EE_ArrayEEUnc = new ArrayList();

            XML_EE_ArrayS11Reel = new ArrayList();
            XML_EE_ArrayS11ReelUnc = new ArrayList();
            XML_EE_ArrayS11Complex = new ArrayList();
            XML_EE_ArrayS11ComplexUnc = new ArrayList();
            XML_EE_ArrayRhoLin = new ArrayList();
            XML_EE_ArrayRhoUnc = new ArrayList();
            XML_EE_ArrayCF = new ArrayList();
            XML_EE_ArrayCFUnc = new ArrayList();
            XML_CF_ArrayFrekans = new ArrayList();

            XML_CF_Array = new ArrayList();
            XML_CF_ArrayCFUnc = new ArrayList();
            XML_CF_ArrayReel = new ArrayList();
            XML_CF_ArrayReelUnc = new ArrayList();
            XML_CF_ArrayComplex = new ArrayList();
            XML_CF_ArrayComplexUnc = new ArrayList();
            XML_CF_YK = new ArrayList();
            XML_CF_YK_Unc = new ArrayList();

            XML_CIS_Olcum_Adım = new ArrayList();
            XML_CIS_ZP = new ArrayList();
            XML_CIS_ZP_Unc = new ArrayList();
            XML_CIS_ICOD = new ArrayList();
            XML_CIS_ICOD_Unc = new ArrayList();
            XML_CIS_OCID = new ArrayList();
            XML_CIS_OCID_Unc = new ArrayList();
        }

        public void SP_ClearData()
        {
            XmlArrayFrekans.Clear();
            XmlArrayS11Reel.Clear();
            XmlArrayS11ReelUnc.Clear();
            XmlArrayS11Complex.Clear();
            XmlArrayS11ComplexUnc.Clear();
            XmlArrayS11Lin.Clear();
            XmlArrayS11LinUnc.Clear();
            XmlArrayS11LinPhase.Clear();
            XmlArrayS11LinPhaseUnc.Clear();
            XmlArrayS11Log.Clear();
            XmlArrayS11LogUnc.Clear();
            XmlArrayS11LogPhase.Clear();
            XmlArrayS11LogPhaseUnc.Clear();
            XmlArrayS11SWR.Clear();
            XmlArrayS11SWRUnc.Clear();
            XmlArrayS12Reel.Clear();
            XmlArrayS12ReelUnc.Clear();
            XmlArrayS12Complex.Clear();
            XmlArrayS12ComplexUnc.Clear();
            XmlArrayS12Lin.Clear();
            XmlArrayS12LinUnc.Clear();
            XmlArrayS12LinPhase.Clear();
            XmlArrayS12LinPhaseUnc.Clear();
            XmlArrayS12Log.Clear();
            XmlArrayS12LogUnc.Clear();
            XmlArrayS12LogPhase.Clear();
            XmlArrayS12LogPhaseUnc.Clear();
            XmlArrayS21Reel.Clear();
            XmlArrayS21ReelUnc.Clear();
            XmlArrayS21Complex.Clear();
            XmlArrayS21ComplexUnc.Clear();
            XmlArrayS21Lin.Clear();
            XmlArrayS21LinUnc.Clear();
            XmlArrayS21LinPhase.Clear();
            XmlArrayS21LinPhaseUnc.Clear();
            XmlArrayS21Log.Clear();
            XmlArrayS21LogUnc.Clear();
            XmlArrayS21LogPhase.Clear();
            XmlArrayS21LogPhaseUnc.Clear();
            XmlArrayS22Reel.Clear();
            XmlArrayS22ReelUnc.Clear();
            XmlArrayS22Complex.Clear();
            XmlArrayS22ComplexUnc.Clear();
            XmlArrayS22Lin.Clear();
            XmlArrayS22LinUnc.Clear();
            XmlArrayS22LinPhase.Clear();
            XmlArrayS22LinPhaseUnc.Clear();
            XmlArrayS22Log.Clear();
            XmlArrayS22LogUnc.Clear();
            XmlArrayS22LogPhase.Clear();
            XmlArrayS22LogPhaseUnc.Clear();
            XmlArrayS22SWR.Clear();
            XmlArrayS22SWRUnc.Clear();
        }
        #endregion

        #region EE 
        public void EE_Data_Xml(string ExcelDosyaYolu, string pageName ,int satır, string sütun)
        {
            int harfIndex = columnName.IndexOf(sütun);
            using (var package = new ExcelPackage(new FileInfo(ExcelDosyaYolu)))
            {


                // Excel'in 13. sayfasındaki veriler
                
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
                        XML_EE_ArrayFrekans.Add(cellValue[i - satır]);
                    }
                }

                // S-Parametre değerlerinin çekimi
                for (int i = satır; i < XML_EE_ArrayFrekans.Count + satır; i++)
                {

                    XML_EE_ArrayEE.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex+1] + i].Value));
                    XML_EE_ArrayEEUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex+2] + i].Value));

                    XML_EE_ArrayS11Reel.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 4] + i].Value));
                    XML_EE_ArrayS11ReelUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 5] + i].Value));
                    XML_EE_ArrayS11Complex.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 6] + i].Value));
                    XML_EE_ArrayS11ComplexUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 7] + i].Value));

                    XML_EE_ArrayRhoLin.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 9] + i].Value));
                    XML_EE_ArrayRhoUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 10] + i].Value));

                    XML_EE_ArrayCF.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 12] + i].Value));
                    XML_EE_ArrayCFUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 13] + i].Value));

                }
            }
        }

        public ArrayList XML_EE_ArrayFrekans { get; set; }


        public ArrayList XML_EE_ArrayEE { get; set; }
        public ArrayList XML_EE_ArrayEEUnc { get; set; }

        public ArrayList XML_EE_ArrayS11Reel { get; set; }
        public ArrayList XML_EE_ArrayS11ReelUnc { get; set; }
        public ArrayList XML_EE_ArrayS11Complex { get; set; }
        public ArrayList XML_EE_ArrayS11ComplexUnc { get; set; }

        public ArrayList XML_EE_ArrayRhoLin { get; set; }
        public ArrayList XML_EE_ArrayRhoUnc { get; set; }

        public ArrayList XML_EE_ArrayCF { get; set; }
        public ArrayList XML_EE_ArrayCFUnc { get; set; }


        public void EE_ClearData()
        {
            XML_EE_ArrayFrekans.Clear();
            XML_EE_ArrayEE.Clear();
            XML_EE_ArrayEEUnc.Clear();
            XML_EE_ArrayS11Reel.Clear();
            XML_EE_ArrayS11ReelUnc.Clear();
            XML_EE_ArrayS11Complex.Clear();
            XML_EE_ArrayS11ComplexUnc.Clear();
            XML_EE_ArrayRhoLin.Clear();
            XML_EE_ArrayRhoUnc.Clear();
            XML_EE_ArrayCF.Clear();
            XML_EE_ArrayCFUnc.Clear();

        }


        #endregion

        #region CF
        public void CF_Data_Xml(string ExcelDosyaYolu, string pageName,int satır, string sütun)
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
                        XML_CF_ArrayFrekans.Add(cellValue[i - satır]);
                    }
                }

                // S-Parametre değerlerinin çekimi
                for (int i = satır; i < XML_CF_ArrayFrekans.Count + satır; i++)
                {

                    XML_CF_Array.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex+1] + i].Value));
                    XML_CF_ArrayCFUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 2] + i].Value));

                    XML_CF_ArrayReel.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 6] + i].Value));
                    XML_CF_ArrayReelUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 7] + i].Value));

                    XML_CF_ArrayComplex.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 8] + i].Value));
                    XML_CF_ArrayComplexUnc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 9] + i].Value));

                    XML_CF_YK.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 10] + i].Value));
                    XML_CF_YK_Unc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 11] + i].Value));



                }
            }
        }

        public ArrayList XML_CF_ArrayFrekans { get; set; }


        public ArrayList XML_CF_Array { get; set; }
        public ArrayList XML_CF_ArrayCFUnc { get; set; }

        public ArrayList XML_CF_ArrayReel { get; set; }
        public ArrayList XML_CF_ArrayReelUnc { get; set; }
        public ArrayList XML_CF_ArrayComplex { get; set; }
        public ArrayList XML_CF_ArrayComplexUnc { get; set; }
        public ArrayList XML_CF_YK { get; set; }
        public ArrayList XML_CF_YK_Unc { get; set; }



        public void CF_ClearData()
        {
            XML_CF_ArrayFrekans.Clear();
            XML_CF_Array.Clear();
            XML_CF_ArrayCFUnc.Clear();
            XML_CF_ArrayReel.Clear();
            XML_CF_ArrayReelUnc.Clear();
            XML_CF_ArrayComplex.Clear();
            XML_CF_ArrayComplexUnc.Clear();
            XML_CF_YK.Clear();
            XML_CF_YK_Unc.Clear();


        }
        #endregion

        #region CIS
        public void CIS_Data_Xml(string ExcelDosyaYolu, string pageName, int satır, string sütun)
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
                        XML_CIS_Olcum_Adım.Add(cellValue[i - satır]);
                    }
                }

                for (int i = satır; i < XML_CIS_Olcum_Adım.Count + satır; i++)
                {



                    XML_CIS_ZP.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex+1] + i].Value));
                    XML_CIS_ZP_Unc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 2] + i].Value));

                    XML_CIS_ICOD.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 3] + i].Value));
                    XML_CIS_ICOD_Unc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 4] + i].Value));

                    XML_CIS_OCID.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 5] + i].Value));
                    XML_CIS_OCID_Unc.Add(Convert.ToDouble(worksheet.Cells[columnName[harfIndex + 6] + i].Value));



                }
            }
        }

        //CIS EXCEL  DEĞERİ (Z POSİTİON)
        public ArrayList XML_CIS_Olcum_Adım { get; set; }

        public ArrayList XML_CIS_ZP { get; set; }
        public ArrayList XML_CIS_ZP_Unc { get; set; }
        public ArrayList XML_CIS_ICOD { get; set; }
        public ArrayList XML_CIS_ICOD_Unc { get; set; }
        public ArrayList XML_CIS_OCID { get; set; }
        public ArrayList XML_CIS_OCID_Unc { get; set; }


        public void CIS_ClearData()
        {
            XML_CIS_Olcum_Adım.Clear();
            XML_CIS_ZP.Clear();
            XML_CIS_ZP_Unc.Clear();
            XML_CIS_ICOD.Clear();
            XML_CIS_ICOD_Unc.Clear();
            XML_CIS_OCID.Clear();
            XML_CIS_OCID_Unc.Clear();

        }

        #endregion
    }
}
