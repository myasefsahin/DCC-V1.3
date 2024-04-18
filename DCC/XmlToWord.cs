using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Charts;
using ExcelToInterface;
using System.Xml;
using System.Collections;
using System.Windows.Forms;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using System.IO;
using System.Xml.XPath;
using DocumentFormat.OpenXml.Presentation;

namespace DCC
{
    public class XmlToWord
    {
        private string dcc = "https://ptb.de/dcc";
        private string si = "https://ptb.de/si";
        public List<bool> dataList = new List<bool>();
        public List<string> headers = new List<string>();
        CreateTable ctb = new CreateTable();
        List<Table> tables = new List<Table>();

        private string orderNo = "/dcc:digitalCalibrationCertificate/dcc:administrativeData/dcc:coreData/dcc:identifications/dcc:identification[@id='orderno']/dcc:value";
        private string itemName = "/dcc:digitalCalibrationCertificate/dcc:administrativeData/dcc:items/dcc:item/dcc:name[@id='itemname']/dcc:content";
        private string itemSerialNumber = "/dcc:digitalCalibrationCertificate/dcc:administrativeData/dcc:items/dcc:item/dcc:identifications/dcc:identification[@id='serialnumber']/dcc:value";

        static void ProcessXmlNodeList(XmlNode node, XmlNamespaceManager nsmgr, ArrayList SvaluesArrays, string SvalueName)
        {
            // XmlNodeList'i seç
            XmlNodeList nodes = node.SelectNodes(SvalueName, nsmgr);

            // Değerleri ayır ve ArrayList'e ekle
            foreach (XmlNode itemNode in nodes)
            {
                string valueString = itemNode.InnerText.Trim(); // quantity tag'ının içeriğini al, başındaki ve sonundaki boşlukları temizle

                string[] values = valueString.Split(' '); // Boşluklara göre değerleri ayır

                foreach (string value in values)
                {
                    string processedValue = value; // İşlenmiş değeri saklamak için bir değişken oluştur

                    // Eğer değerde E varsa işlem yap
                    if (value.Contains("E"))
                    {
                        // E karakterinden önceki kısmı al, E karakterini ve sonrasını sil
                        int eIndex = value.IndexOf("E");
                        processedValue = value.Substring(0, eIndex);
                    }

                    decimal parsedValue;
                    if (decimal.TryParse(processedValue, out parsedValue))
                    {
                        SvaluesArrays.Add(parsedValue); // Değeri ArrayList'e ekle
                    }
                    else if (!string.IsNullOrWhiteSpace(processedValue)) // Eğer değer null ya da boş değilse
                    {
                        SvaluesArrays.Add(processedValue); // Değeri ArrayList'e ekle (string olarak)
                    }
                }
            }
        }
        public void Try(string filePath)
        {
            try
            {
                // XML dosyasının yolu
                string xmlDosyaYolu = filePath;

                // XML yapısını yükleyeceğimiz XmlDocument oluşturun
                XmlDocument xmlDoc = new XmlDocument();
                ExcelData exd = new ExcelData();


                // XML dosyasını yükleyin
                xmlDoc.Load(xmlDosyaYolu);

                // XML namespace'lerini oluşturun
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
                nsmgr.AddNamespace("si", "https://ptb.de/si");

                XmlNode orderNoNode = xmlDoc.SelectSingleNode(orderNo, nsmgr);
                XmlNode itemNameNode = xmlDoc.SelectSingleNode(itemName, nsmgr);
                XmlNode itemSerialNoNode = xmlDoc.SelectSingleNode(itemSerialNumber, nsmgr);
                // "result" tag'lerini içeren düğümleri seçin
                XmlNodeList resultNodes = xmlDoc.SelectNodes("/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result", nsmgr);
                foreach (XmlNode resultNode in resultNodes)
                {
                    XElement resultElement = XElement.Parse(resultNode.OuterXml);
                    List<bool> boolList = SelectFilledColumns(resultElement);
                    #region Tanımlar
                    #region Frekanslar
                    string frequencyS = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_sp']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequencyEE = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ee']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequencyCF = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_cf']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequencyCIS = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_cis']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequencyNoise = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_noise']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RF_Dif_Freq1 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RF_Diff_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RF_Dif_Freq2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RF_Diff_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RF_Dif_Freq3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RF_Diff_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RF_Dif_Freq4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RF_Diff_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RF_Gain_Freq= "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Gain_input_nom_freq']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequency_ARFP_t1 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ARFP_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequency_ARFP_t2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ARFP_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequency_ARFP_t3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ARFP_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequency_ARFP_t4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ARFP_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequency_ARFP_t5 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ARFP_t5']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequency_ARFP_t6 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ARFP_t6']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequency_ARFP_t7 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ARFP_t7']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequency_ARFP_t8 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ARFP_t8']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequency_ARFP_t9 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ARFP_t9']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequency_ARFP_t10 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ARFP_t10']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string frequency_ARFP_t11= "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='frequency_ARFP_t11']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";

                    #endregion
                    #region S_Param
                    string s11reel = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11Reel']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11reelUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11ReelUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11Imag = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11Imag']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11ImagUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11ImagUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11Lin = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11Lin']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11LinUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11LinUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11LinPhase = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11Phase']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11LinPhaseUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11PhaseUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11Log = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11Log']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11LogUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11LogUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11LogPhase = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11LogPhase']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11LogPhaseUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11LogPhaseUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11SWR = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11swr']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s11SWRUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters11swrUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12reel = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12Reel']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12reelUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12ReelUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12Imag = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12Imag']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12ImagUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12ImagUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12Lin = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12Lin']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12LinUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12LinUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12LinPhase = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12Phase']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12LinPhaseUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12PhaseUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12Log = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12Log']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12LogUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12LogUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12LogPhase = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12LogPhase']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s12LogPhaseUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters12LogPhaseUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21reel = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21Reel']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21reelUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21ReelUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21Imag = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21Imag']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21ImagUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21ImagUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21Lin = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21Lin']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21LinUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21LinUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21LinPhase = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21Phase']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21LinPhaseUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21PhaseUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21Log = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21Log']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21LogUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21LogUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21LogPhase = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21LogPhase']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s21LogPhaseUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters21LogPhaseUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22reel = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22Reel']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22reelUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22ReelUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22Imag = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22Imag']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22ImagUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22ImagUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22Lin = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22Lin']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22LinUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22LinUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22LinPhase = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22Phase']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22LinPhaseUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22PhaseUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22Log = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22Log']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22LogUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22LogUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22LogPhase = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22LogPhase']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22LogPhaseUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22LogPhaseUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22SWR = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22swr']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string s22SWRUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='s_parameters22swrUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    #endregion
                    #region EE
                    string EffiencyEEEE = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Effective Effiency EE-EE']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string EffiencyEEEE_Unc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Effective Effiency EE-EE_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string EffiencyEE_S11Reel = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Effective Effiency EE-Reel']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string EffiencyEE_S11ReelUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Effective Effiency EE-Reel_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string EffiencyEE_S11Imag = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Effective Effiency EE-Imaginer']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string EffiencyEE_S11ImagUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Effective Effiency EE-Imaginer_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string EffiencyRHO_EERho = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Effective Effiency EE-Rho']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string EffiencyRHO_EERhoUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Effective Effiency EE-Rho_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string EffiencyEE_CFEE_CF = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Effective Effiency EE-Cal_Factor']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string EffiencyEE_CFEE_CFUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Effective Effiency EE-Cal_Factor_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    #endregion
                    #region CF
                    string CF_Cal_Factor = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calibration Factor CF-Cal_Factor']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CF_Cal_Factor_Unc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calibration Factor CF-Cal_Factor_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CF_Reel = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calibration Factor CF-Reel']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CF_Reel_Unc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calibration Factor CF-Reel_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CF_Imaginer = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calibration Factor CF-Imaginer']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CF_Imaginer_Unc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calibration Factor CF-Imaginer_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CF_ReflectionCof = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calibration Factor CF-ReflectionCof']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CF_ReflectionCof_Unc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calibration Factor CF-ReflectionCof_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";

                    string CIS_Z_Position = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calculable Impedance Standard CIS-Z-Position']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CIS_Z_Position_Unc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calculable Impedance Standard CIS-Z-PositionUnc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CIS_ICOD = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calculable Impedance Standard CIS-ICOD']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CIS_ICOD_Unc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calculable Impedance Standard CIS-ICOD_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CIS_OCID = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calculable Impedance Standard CIS-OCID']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string CIS_OCID_Unc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Calculable Impedance Standard CIS-OCID_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    #endregion
                    #region Noise
                    string NoiseENR = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_ENR']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string NoiseENRUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_ENR_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string NoiseDCONRCLin = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_DC_ON_Lin']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string NoiseDCONUpLimit = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_DC_ON_Limit']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string NoiseDCONRCLinUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_DC_ON_Lin_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string NoiseDCONRCPhase = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_DC_ON_RC_Phase']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string NoiseDCONRCPhaseUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_DC_ON_RC_Phase_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string NoiseDCOFFRCLin = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_DC_OFF_Lin']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string DCOFFUpLimit = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_DC_OFF_Limit']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string NoiseDCOFFRCLinUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_DC_OFF_Lin_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string NoiseDCOFFRCPhase = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_DC_OFF_RC_Phase']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string NoiseDCOFFRCPhaseUnc = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Noise_DC_OFF_RC_Phase_Unc']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    #endregion
                    #region RF_Diff
                    string RFD_IndıcatorVal = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_IndıcatorVal_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_lowerLimit = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_lowerLimit_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_measuredVal = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_measuredVal_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_measureDiff = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_measureDiff_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_upperLimit = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_upperLimit_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_uncertainty = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_uncertainty_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";

                    string RFD_NomPowlvl2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_NomPowlvl_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_measuredVal2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_measuredVal_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_lowerLimit2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_lowerLimit_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_NomPowlvlDiff2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_NomPowlvlDiff_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_upperLimit2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_upperLimit_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_uncertainty2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_uncertainty_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";

                    string RFD_Nom_pow3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_Nom_pow_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_lowerLimit3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_lowerLimit_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_measuredVal3= "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_measuredVal_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_upperLimit3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_upperLimit_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_difference3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_difference_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_uncertainty3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_uncertainty_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";

                    string RFD_MinPowLevel4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_MinPowLevel_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_MaxPowLevel4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_MaxPowLevel_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_LowerLimit4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_LowerLimit_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_difference4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_difference_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_upper_limit4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_upper_limit_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFD_uncertainty4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFD_uncertainty_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";

                    #endregion
                    #region RF_Gain
                    string RFG_Input_Pow1 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFG_Input_Pow1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFG_Unc1 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFG_Unc1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Biggest_gain = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Biggest_gain']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFG_lowest_Gain = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFG_lowest_Gain']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFG_Flatness = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFG_Flatness']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Gain_diff_input_100KHz = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Gain_diff_input_100KHz']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFG_Input_Pow2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFG_Input_Pow2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFG_Unc2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFG_Unc2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Gain_diff_input_1GHz = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Gain_diff_input_1GHz']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFG_Input_Pow3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFG_Input_Pow3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string RFG_Unc3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='RFG_Unc3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    #endregion
                    #region RF_Absolute
                    string Abs_RF_Power_Output_Power_t1 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Output_Power_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Measured_Power_t1 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Measured_Power_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Lower_limit_t1 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Lower_limit_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Deflection_t1 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Deflection_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Upper_Limit_t1 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Upper_Limit_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Uncertainty_t1 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Uncertainty_t1']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Output_Power_t2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Output_Power_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Measured_Power_t2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Measured_Power_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Lower_limit_t2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Lower_limit_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_difference_t2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_difference_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Upper_Limit_t2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Upper_Limit_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Uncertainty_t2 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Uncertainty_t2']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Output_Power_t3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Output_Power_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Measured_Attenuation_t3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Measured_Attenuation_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_lower_Limit_t3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_lower_Limit_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Attenuation_Error_t3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Attenuation_Error_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Upper_Limit_t3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Upper_Limit_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Uncertainty_t3 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Uncertainty_t3']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Level_t4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Level_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Measured_Value_t4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Measured_Value_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Maximum_t4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Maximum_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Uncertainty_t4 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Uncertainty_t4']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Level_t5 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Level_t5']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Measured_Value_t5 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Measured_Value_t5']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Maximum_t5 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Maximum_t5']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Uncertainty_t5 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Uncertainty_t5']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Level_t6 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Level_t6']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Measured_Value_t6 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Measured_Value_t6']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Maximum_t6 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Maximum_t6']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Uncertainty_t6 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Uncertainty_t6']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Output_Power_t7 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Output_Power_t7']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Measured_Power_t7 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Measured_Power_t7']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Lower_Limit_t7 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Lower_Limit_t7']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Deflection_t7 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Deflection_t7']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Upper_Limit_t7 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Upper_Limit_t7']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Uncertainty_t7 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Uncertainty_t7']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Output_Power_t8 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Output_Power_t8']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Measured_Value_t8 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Measured_Value_t8']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Lower_Limit_t8 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Lower_Limit_t8']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Difference_t8 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Difference_t8']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Upper_Limit_t8 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Upper_Limit_t8']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Uncertainty_t8 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Uncertainty_t8']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Level_t9 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Level_t9']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Measured_Value_t9 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Measured_Value_t9']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Upper_Limit_t9 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Upper_Limit_t9']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Uncertainty_t9 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Uncertainty_t9']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Level_t10 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Level_t10']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Measured_Value_t10 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Measured_Value_t10']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Upper_Limit_t10 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Upper_Limit_t10']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Uncertainty_t10 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Uncertainty_t10']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Level_t11 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Level_t11']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Measured_Value_t11 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Measured_Value_t11']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Upper_Limit_t11 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Upper_Limit_t11']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";
                    string Abs_RF_Power_Uncertainty_t11 = "/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results/dcc:result/dcc:data/dcc:list/dcc:quantity[@refType='Abs_RF_Power_Uncertainty_t11']/si:hybrid/si:realListXMLList/si:valueXMLList/text()";

                    #endregion
                    #endregion
                    #region verileri Excell Data Nesnelerine tanımlama 
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayFrekansSParam, frequencyS);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayFrekansEE, frequencyEE);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayFrekansCF, frequencyCF);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayFrekansCIS, frequencyCIS);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayFrekansNoise, frequencyNoise);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T1_Frekans, RF_Dif_Freq1);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T2_Frekans, RF_Dif_Freq2);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T3_Frekans, RF_Dif_Freq3);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T4_Frekans, RF_Dif_Freq4);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T1_Frekans, RF_Gain_Freq);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T2_EnBuyukKazanc, Biggest_gain);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T3_Nom_Giris_Gucu, Gain_diff_input_100KHz);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T4_Nom_Giris_Gucu, Gain_diff_input_1GHz);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T1_Frekans, frequency_ARFP_t1);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T2_Frekans, frequency_ARFP_t2);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T3_Frekans, frequency_ARFP_t3);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T4_T5_T6_frekans, frequency_ARFP_t4);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T7_Frekans, frequency_ARFP_t7);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T8_Frekans, frequency_ARFP_t8);
                    ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T9_T10_T11_frekans, frequency_ARFP_t9);

                    if (boolList[0] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11Reel, s11reel);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11ReelUnc, s11reelUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11Complex, s11Imag);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11ComplexUnc, s11ImagUnc);
                    }
                    if (boolList[1] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11Lin, s11Lin);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11LinUnc, s11LinUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11LinPhase, s11LinPhase);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11LinPhaseUnc, s11LinPhaseUnc);
                    }
                    if (boolList[2] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11Log, s11Log);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11LogUnc, s11LogUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11LogPhase, s11LogPhase);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11LogPhaseUnc, s11LogPhaseUnc);
                    }
                    if (boolList[3] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11SWR, s11SWR);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS11SWRUnc, s11SWRUnc);
                    }
                    if (boolList[4] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12Reel, s12reel);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12ReelUnc, s12reelUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12Complex, s12Imag);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12ComplexUnc, s12ImagUnc);
                    }
                    if (boolList[5] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12Lin, s12Lin);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12LinUnc, s12LinUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12LinPhase, s12LinPhase);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12LinPhaseUnc, s12LinPhaseUnc);
                    }
                    if (boolList[6] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12Log, s12Log);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12LogUnc, s12LogUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12LogPhase, s12LogPhase);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS12LogPhaseUnc, s12LogPhaseUnc);
                    }
                    if (boolList[7] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21Reel, s21reel);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21ReelUnc, s21reelUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21Complex, s21Imag);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21ComplexUnc, s21ImagUnc);
                    }
                    if (boolList[8] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21Lin, s21Lin);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21LinUnc, s21LinUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21LinPhase, s21LinPhase);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21LinPhaseUnc, s21LinPhaseUnc);
                    }
                    if (boolList[9] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21Log, s21Log);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21LogUnc, s21LogUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21LogPhase, s21LogPhase);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS21LogPhaseUnc, s21LogPhaseUnc);
                    }
                    if (boolList[10] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22Reel, s22reel);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22ReelUnc, s22reelUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22Complex, s22Imag);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22ComplexUnc, s22ImagUnc);
                    }
                    if (boolList[11] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22Lin, s22Lin);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22LinUnc, s22LinUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22LinPhase, s22LinPhase);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22LinPhaseUnc, s22LinPhaseUnc);
                    }
                    if (boolList[12] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22Log, s22Log);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22LogUnc, s22LogUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22LogPhase, s22LogPhase);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22LogPhaseUnc, s22LogPhaseUnc);

                    }
                    if (boolList[13] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22SWR, s22SWR);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayS22SWRUnc, s22SWRUnc);
                    }
                    if (boolList[14] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayEffiencyEEEE, EffiencyEEEE);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayEffiencyEEEEUnc, EffiencyEEEE_Unc);
                    }
                    if (boolList[15] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayEffiencyEE_S11Reel, EffiencyEE_S11Reel);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayEffiencyEE_S11ReelUnc, EffiencyEE_S11ReelUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayEffiencyEE_S11Imag, EffiencyEE_S11Imag);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayEffiencyEE_S11ImagUnc, EffiencyEE_S11ImagUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayEffiencyRHO_EERho, EffiencyRHO_EERho);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayEffiencyRHO_EERhoUnc, EffiencyRHO_EERhoUnc);
                    }
                    if (boolList[16] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayEffiencyEE_CFEE_CF, EffiencyEE_CFEE_CF);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayEffiencyEE_CFEE_CFUnc, EffiencyEE_CFEE_CFUnc);

                    }
                    if (boolList[17] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCF_Cal_Factor, CF_Cal_Factor);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCF_Cal_Factor_Unc, CF_Cal_Factor_Unc);
                    }
                    if (boolList[18] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCF_Imaginer, CF_Imaginer);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCF_Imaginer_Unc, CF_Imaginer_Unc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCF_Reel, CF_Reel);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCF_Reel_Unc, CF_Reel_Unc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCF_ReflectionCof, CF_ReflectionCof);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCF_ReflectionCof_Unc, CF_ReflectionCof_Unc);
                    }
                    if (boolList[19] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCIS_Z_Position, CIS_Z_Position);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCIS_Z_Position_Unc, CIS_Z_Position_Unc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCIS_ICOD, CIS_ICOD);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCIS_ICOD_Unc, CIS_ICOD_Unc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCIS_OCID, CIS_OCID);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayCIS_OCID_Unc, CIS_OCID_Unc);

                    }
                    if (boolList[20] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayNoiseENR, NoiseENR);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayNoiseENRUnc, NoiseENRUnc);
                    }
                    if (boolList[21] == true)
                    {

                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayNoiseDCONRCLin, NoiseDCONRCLin);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayNoiseDCONUpLimit, NoiseDCONUpLimit);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayNoiseDCONRCLinUnc, NoiseDCONRCLinUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayNoiseDCONRCPhase, NoiseDCONRCPhase);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayNoiseDCONRCPhaseUnc, NoiseDCONRCPhaseUnc);
                    }
                    if (boolList[22] == true)
                    {

                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayNoiseDCOFFRCLin, NoiseDCOFFRCLin);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayDCOFFUpLimit, DCOFFUpLimit);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayNoiseDCOFFRCLinUnc, NoiseDCOFFRCLinUnc);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayNoiseDCOFFRCPhase, NoiseDCOFFRCPhase);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayNoiseDCOFFRCPhaseUnc, NoiseDCOFFRCPhaseUnc);
                    }
                    if (boolList[23] == true)
                    {

                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T1_GostergeDegeri, RFD_IndıcatorVal);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T1_AltSınır, RFD_lowerLimit);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T1_OlculenDeger, RFD_measuredVal);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T1_OlculenFark, RFD_measureDiff);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T1_UstSınır, RFD_upperLimit);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T1_Belirsizlik, RFD_uncertainty);
                    }
                    if (boolList[24] == true)
                    {

                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T2_Nom_Guc_Lvl, RFD_NomPowlvl2);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T2_OlculenDeger, RFD_measuredVal2);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T2_AltSınır, RFD_lowerLimit2);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T2_Nom_Guc_Lvl_fark, RFD_NomPowlvlDiff2);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T2_UstSınır, RFD_upperLimit2);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T2_Belirsizlik, RFD_uncertainty2);
                    }
                    if (boolList[25] == true)
                    {

                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T3_NominalGuc, RFD_Nom_pow3);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T3_AltSınır, RFD_lowerLimit3);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T3_OlculenDeger, RFD_measuredVal3);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T3_UstSınır, RFD_upperLimit3);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T3_Fark, RFD_difference3);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T3_Belirsizlik, RFD_uncertainty3);
                    }
                    if (boolList[26] == true)
                    {

                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T4_Min_Guc_lvl , RFD_MinPowLevel4);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T4_Max_Guc_lvl , RFD_MaxPowLevel4);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T4_AltSınır , RFD_LowerLimit4);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T4_Fark, RFD_difference4);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T4_UstSınır, RFD_upper_limit4);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFD_T4_Belirsizlik, RFD_uncertainty4);
                    }
                    if (boolList[27] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T1_GirisGucu, RFG_Input_Pow1);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T1_Belirsizlik, RFG_Unc1);
                    }
                    if (boolList[28] == true)
                    {
                     
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T2_EnKucukKazanc, RFG_lowest_Gain);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T2_Flatness, RFG_Flatness);
                    }
                    if (boolList[29] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T3_Kazanc, Gain_diff_input_100KHz);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T3_Belirsizlik, RFG_Unc2);
                    }
                    if (boolList[30] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T4_Kazanc, Gain_diff_input_1GHz);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayRFG_T4_Belirsizlik, RFG_Unc3);
                    }
                    if (boolList[31] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T1_Cıkıs_Gücü, Abs_RF_Power_Output_Power_t1);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T1_Olculen_Güc, Abs_RF_Power_Measured_Power_t1);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T1_AltSınır, Abs_RF_Power_Lower_limit_t1);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T1_Sapma, Abs_RF_Power_Deflection_t1);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T1_ÜstSınır, Abs_RF_Power_Upper_Limit_t1);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T1_Belirsizlik, Abs_RF_Power_Uncertainty_t1);
                    }
                    if (boolList[32] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T2_Cıkıs_Gücü, Abs_RF_Power_Output_Power_t2);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T2_OlculenDeger, Abs_RF_Power_Measured_Power_t2);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T2_AltSınır, Abs_RF_Power_Lower_limit_t2);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T2_Fark, Abs_RF_Power_difference_t2);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T2_ÜstSınır, Abs_RF_Power_Upper_Limit_t2);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T2_Belirsizlik, Abs_RF_Power_Uncertainty_t2);
                    }
                    if (boolList[33] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T3_Cıkıs_Gücü, Abs_RF_Power_Output_Power_t3);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T3_OlculenZayıflatma, Abs_RF_Power_Measured_Attenuation_t3);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T3_AltSınır, Abs_RF_Power_lower_Limit_t3);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T3_Zayıflatma, Abs_RF_Power_Attenuation_Error_t3);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T3_ÜstSınır, Abs_RF_Power_Upper_Limit_t3);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T3_Belirsizlik, Abs_RF_Power_Uncertainty_t3);
                    }
                    if (boolList[34] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T4_SWR_Seviye, Abs_RF_Power_Level_t4);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T4_SWR_OlculenDeger, Abs_RF_Power_Measured_Value_t4);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T4_SWR_MaksimumDeger, Abs_RF_Power_Maximum_t4);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T4_SWR_Belirsizlik, Abs_RF_Power_Uncertainty_t4);
                    }
                    if (boolList[35] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T5_SWR_Seviye, Abs_RF_Power_Level_t5);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T5_SWR_OlculenDeger, Abs_RF_Power_Measured_Value_t5);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T5_SWR_MaksimumDeger, Abs_RF_Power_Maximum_t5);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T5_SWR_Belirsizlik, Abs_RF_Power_Uncertainty_t5);
                    }
                    if (boolList[36] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T6_SWR_Seviye, Abs_RF_Power_Level_t6);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T6_SWR_OlculenDeger, Abs_RF_Power_Measured_Value_t6);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T6_SWR_MaksimumDeger, Abs_RF_Power_Maximum_t6);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T6_SWR_Belirsizlik, Abs_RF_Power_Uncertainty_t6);
                    }
                    if (boolList[37] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T7_Cıkıs_Gücü, Abs_RF_Power_Output_Power_t7);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T7_OlculenGuc, Abs_RF_Power_Measured_Power_t7);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T7_AltSınır, Abs_RF_Power_Lower_Limit_t7);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T7_Sapma, Abs_RF_Power_Deflection_t7);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T7_ÜstSınır, Abs_RF_Power_Upper_Limit_t7);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T7_Belirsizlik, Abs_RF_Power_Uncertainty_t7);
                    }
                    if (boolList[38] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T8_Cıkıs_Gücü, Abs_RF_Power_Output_Power_t8);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T8_OlculenDeger, Abs_RF_Power_Measured_Value_t8);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T8_AltSınır, Abs_RF_Power_Lower_Limit_t8);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T8_Fark, Abs_RF_Power_Difference_t8);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T8_ÜstSınır, Abs_RF_Power_Upper_Limit_t8);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T8_Belirsizlik, Abs_RF_Power_Uncertainty_t8);
                    }
                    if (boolList[39] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T9_SWR_Seviye, Abs_RF_Power_Level_t9);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T9_SWR_OlculenDeger, Abs_RF_Power_Measured_Value_t9);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T9_SWR_MaksimumDeger, Abs_RF_Power_Upper_Limit_t9);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T9_SWR_Belirsizlik, Abs_RF_Power_Uncertainty_t9);
                    }
                    if (boolList[40] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T10_SWR_Seviye, Abs_RF_Power_Level_t10);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T10_SWR_OlculenDeger, Abs_RF_Power_Measured_Value_t10);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T10_SWR_MaksimumDeger, Abs_RF_Power_Upper_Limit_t10);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T10_SWR_Belirsizlik, Abs_RF_Power_Uncertainty_t10);
                    }
                    if (boolList[41] == true)
                    {
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T11_SWR_Seviye, Abs_RF_Power_Level_t11);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T11_SWR_OlculenDeger, Abs_RF_Power_Measured_Value_t11);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T11_SWR_MaksimumDeger, Abs_RF_Power_Upper_Limit_t11);
                        ProcessXmlNodeList(resultNode, nsmgr, exd.ArrayARFP_T11_SWR_Belirsizlik, Abs_RF_Power_Uncertainty_t11);
                    }




                    #endregion
                    #region Verilerin Formatlanması ve Dizilere Atanması 
                    #region Diziler
                    ArrayList arrays11reel = new ArrayList();
                    ArrayList arrays11reelunc = new ArrayList();
                    ArrayList arrays11complex = new ArrayList();
                    ArrayList arrays11complexunc = new ArrayList();
                    ArrayList arrays11lin = new ArrayList();
                    ArrayList arrays11linunc = new ArrayList();
                    ArrayList arrays11linphase = new ArrayList();
                    ArrayList arrays11linphaseunc = new ArrayList();
                    ArrayList arrays11log = new ArrayList();
                    ArrayList arrays11logunc = new ArrayList();
                    ArrayList arrays11logphase = new ArrayList();
                    ArrayList arrays11logphaseunc = new ArrayList();
                    ArrayList arrays11swr = new ArrayList();
                    ArrayList arrays11swrunc = new ArrayList();
                    ArrayList arrays12reel = new ArrayList();
                    ArrayList arrays12reelunc = new ArrayList();
                    ArrayList arrays12complex = new ArrayList();
                    ArrayList arrays12complexunc = new ArrayList();
                    ArrayList arrays12lin = new ArrayList();
                    ArrayList arrays12linunc = new ArrayList();
                    ArrayList arrays12linphase = new ArrayList();
                    ArrayList arrays12linphaseunc = new ArrayList();
                    ArrayList arrays12log = new ArrayList();
                    ArrayList arrays12logunc = new ArrayList();
                    ArrayList arrays12logphase = new ArrayList();
                    ArrayList arrays12logphaseunc = new ArrayList();
                    ArrayList arrays21reel = new ArrayList();
                    ArrayList arrays21reelunc = new ArrayList();
                    ArrayList arrays21complex = new ArrayList();
                    ArrayList arrays21complexunc = new ArrayList();
                    ArrayList arrays21lin = new ArrayList();
                    ArrayList arrays21linunc = new ArrayList();
                    ArrayList arrays21linphase = new ArrayList();
                    ArrayList arrays21linphaseunc = new ArrayList();
                    ArrayList arrays21log = new ArrayList();
                    ArrayList arrays21logunc = new ArrayList();
                    ArrayList arrays21logphase = new ArrayList();
                    ArrayList arrays21logphaseunc = new ArrayList();
                    ArrayList arrays22reel = new ArrayList();
                    ArrayList arrays22reelunc = new ArrayList();
                    ArrayList arrays22complex = new ArrayList();
                    ArrayList arrays22complexunc = new ArrayList();
                    ArrayList arrays22lin = new ArrayList();
                    ArrayList arrays22linunc = new ArrayList();
                    ArrayList arrays22linphase = new ArrayList();
                    ArrayList arrays22linphaseunc = new ArrayList();
                    ArrayList arrays22log = new ArrayList();
                    ArrayList arrays22logunc = new ArrayList();
                    ArrayList arrays22logphase = new ArrayList();
                    ArrayList arrays22logphaseunc = new ArrayList();
                    ArrayList arrays22swr = new ArrayList();
                    ArrayList arrays22swrunc = new ArrayList();
                    ArrayList arraysEffiencyEEEE = new ArrayList();
                    ArrayList arraysEffiencyEEEEunc = new ArrayList();
                    ArrayList arraysEffiencyEE_S11Reel = new ArrayList();
                    ArrayList arraysEffiencyEE_S11Reelunc = new ArrayList();
                    ArrayList arraysEffiencyEE_S11Imag = new ArrayList();
                    ArrayList arraysEffiencyEE_S11Imagunc = new ArrayList();
                    ArrayList arraysEffiencyRHO_EERho = new ArrayList();
                    ArrayList arraysEffiencyRHO_EERhounc = new ArrayList();
                    ArrayList arraysEffiencyEE_CFEE_CF = new ArrayList();
                    ArrayList arraysEffiencyEE_CFEE_CFunc = new ArrayList();
                    ArrayList arrayCF_Cal_Factor = new ArrayList();
                    ArrayList arrayCF_Cal_Factor_Unc = new ArrayList();
                    ArrayList arrayCF_Reel = new ArrayList();
                    ArrayList arrayCF_Reel_Unc = new ArrayList();
                    ArrayList arrayCF_Imaginer = new ArrayList();
                    ArrayList arrayCF_Imaginer_Unc = new ArrayList();
                    ArrayList arrayCF_ReflectionCof = new ArrayList();
                    ArrayList arrayCF_ReflectionCof_Unc = new ArrayList();
                    ArrayList arrayCIS_Z_Position = new ArrayList();
                    ArrayList arrayCIS_Z_Position_Unc = new ArrayList();
                    ArrayList arrayCIS_ICOD = new ArrayList();
                    ArrayList arrayCIS_ICOD_Unc = new ArrayList();
                    ArrayList arrayCIS_OCID = new ArrayList();
                    ArrayList arrayCIS_OCID_Unc = new ArrayList();

                    ArrayList arrayNoiseENR = new ArrayList();
                    ArrayList arrayNoiseENRUnc = new ArrayList();
                    ArrayList arrayNoiseDCONRCLin = new ArrayList();
                    ArrayList arrayNoiseDCONUpLimit = new ArrayList();
                    ArrayList arrayNoiseDCONRCLinUnc = new ArrayList();
                    ArrayList arrayNoiseDCONRCPhase = new ArrayList();
                    ArrayList arrayNoiseDCONRCPhaseUnc = new ArrayList();
                    ArrayList arrayNoiseDCOFFRCLin = new ArrayList();
                    ArrayList arrayDCOFFUpLimit = new ArrayList();
                    ArrayList arrayNoiseDCOFFRCLinUnc = new ArrayList();
                    ArrayList arrayNoiseDCOFFRCPhase = new ArrayList();
                    ArrayList arrayNoiseDCOFFRCPhaseUnc = new ArrayList();

                    ArrayList arrayRFD_T1_Frekans = new ArrayList();
                    ArrayList arrayRFD_T1_GostergeDegeri = new ArrayList();
                    ArrayList arrayRFD_T1_AltSınır = new ArrayList();
                    ArrayList arrayRFD_T1_OlculenDeger = new ArrayList();
                    ArrayList arrayRFD_T1_OlculenFark = new ArrayList();
                    ArrayList arrayRFD_T1_UstSınır = new ArrayList();
                    ArrayList arrayRFD_T1_Belirsizlik = new ArrayList();
                    ArrayList arrayRFD_T1_Belirsizlik2 = new ArrayList();
                    ArrayList arrayRFD_T1_Belirsizlik3 = new ArrayList();

                    ArrayList arrayRFD_T2_Frekans = new ArrayList();
                    ArrayList arrayRFD_T2_Nom_Guc_Lvl = new ArrayList();
                    ArrayList arrayRFD_T2_OlculenDeger = new ArrayList();
                    ArrayList arrayRFD_T2_AltSınır = new ArrayList();
                    ArrayList arrayRFD_T2_Nom_Guc_Lvl_fark = new ArrayList();
                    ArrayList arrayRFD_T2_UstSınır = new ArrayList();
                    ArrayList arrayRFD_T2_Belirsizlik = new ArrayList();
                    ArrayList arrayRFD_T2_Belirsizlik2 = new ArrayList();

                    ArrayList arrayRFD_T3_Frekans = new ArrayList();
                    ArrayList arrayRFD_T3_NominalGuc = new ArrayList();
                    ArrayList arrayRFD_T3_AltSınır = new ArrayList();
                    ArrayList arrayRFD_T3_OlculenDeger = new ArrayList();
                    ArrayList arrayRFD_T3_UstSınır = new ArrayList();
                    ArrayList arrayRFD_T3_Fark = new ArrayList();
                    ArrayList arrayRFD_T3_Belirsizlik = new ArrayList();
                    ArrayList arrayRFD_T3_Belirsizlik2 = new ArrayList();

                    ArrayList arrayRFD_T4_Min_Guc_lvl = new ArrayList();
                    ArrayList arrayRFD_T4_Max_Guc_lvl = new ArrayList();
                    ArrayList arrayRFD_T4_Frekans = new ArrayList();
                    ArrayList arrayRFD_T4_AltSınır = new ArrayList();
                    ArrayList arrayRFD_T4_Fark = new ArrayList();
                    ArrayList arrayRFD_T4_UstSınır = new ArrayList();
                    ArrayList arrayRFD_T4_Belirsizlik = new ArrayList();
           

                    ArrayList arrayRFG_T1_Frekans = new ArrayList();
                    ArrayList arrayRFG_T1_GirisGucu = new ArrayList();
                    ArrayList arrayRFG_T1_Belirsizlik = new ArrayList();

                    ArrayList arrayRFG_T2_EnBuyukKazanc = new ArrayList();
                    ArrayList arrayRFG_T2_EnKucukKazanc = new ArrayList();
                    ArrayList arrayRFG_T2_Flatness = new ArrayList();

                    ArrayList arrayRFG_T3_Nom_Giris_Gucu = new ArrayList();
                    ArrayList arrayRFG_T3_Kazanc = new ArrayList();
                    ArrayList arrayRFG_T3_Belirsizlik = new ArrayList();

                    ArrayList arrayRFG_T4_Nom_Giris_Gucu = new ArrayList();
                    ArrayList arrayRFG_T4_Kazanc = new ArrayList();
                    ArrayList arrayRFG_T4_Belirsizlik = new ArrayList();

                    ArrayList arrayARFP_T1_Frekans = new ArrayList();
                    ArrayList arrayARFP_T1_Cıkıs_Gücü = new ArrayList();
                    ArrayList arrayARFP_T1_Olculen_Güc = new ArrayList();
                    ArrayList arrayARFP_T1_AltSınır = new ArrayList();
                    ArrayList arrayARFP_T1_Sapma = new ArrayList();
                    ArrayList arrayARFP_T1_ÜstSınır = new ArrayList();
                    ArrayList arrayARFP_T1_Belirsizlik = new ArrayList();
                    ArrayList arrayARFP_T1_Belirsizlik2 = new ArrayList();
                    ArrayList arrayARFP_T1_Belirsizlik3 = new ArrayList();

                    ArrayList arrayARFP_T2_Frekans = new ArrayList();
                    ArrayList arrayARFP_T2_Cıkıs_Gücü = new ArrayList();
                    ArrayList arrayARFP_T2_OlculenDeger = new ArrayList();
                    ArrayList arrayARFP_T2_AltSınır = new ArrayList();
                    ArrayList arrayARFP_T2_Fark = new ArrayList();
                    ArrayList arrayARFP_T2_ÜstSınır = new ArrayList();
                    ArrayList arrayARFP_T2_Belirsizlik = new ArrayList();
                    ArrayList arrayARFP_T2_Belirsizlik2 = new ArrayList();

                    ArrayList arrayARFP_T3_Frekans = new ArrayList();
                    ArrayList arrayARFP_T3_Cıkıs_Gücü = new ArrayList();
                    ArrayList arrayARFP_T3_OlculenZayıflatma = new ArrayList();
                    ArrayList arrayARFP_T3_AltSınır = new ArrayList();
                    ArrayList arrayARFP_T3_Zayıflatma = new ArrayList();
                    ArrayList arrayARFP_T3_ÜstSınır = new ArrayList();
                    ArrayList arrayARFP_T3_Belirsizlik = new ArrayList();
                    ArrayList arrayARFP_T3_Belirsizlik2 = new ArrayList();

                    ArrayList arrayARFP_T4_T5_T6_frekans = new ArrayList();
                    ArrayList arrayARFP_T4_SWR_Seviye = new ArrayList();
                    ArrayList arrayARFP_T4_SWR_OlculenDeger = new ArrayList();
                    ArrayList arrayARFP_T4_SWR_MaksimumDeger = new ArrayList();
                    ArrayList arrayARFP_T4_SWR_Belirsizlik = new ArrayList();

                    ArrayList arrayARFP_T5_SWR_Seviye = new ArrayList();
                    ArrayList arrayARFP_T5_SWR_OlculenDeger = new ArrayList();
                    ArrayList arrayARFP_T5_SWR_MaksimumDeger = new ArrayList();
                    ArrayList arrayARFP_T5_SWR_Belirsizlik = new ArrayList();

                    ArrayList arrayARFP_T6_SWR_Seviye = new ArrayList();
                    ArrayList arrayARFP_T6_SWR_OlculenDeger = new ArrayList();
                    ArrayList arrayARFP_T6_SWR_MaksimumDeger = new ArrayList();
                    ArrayList arrayARFP_T6_SWR_Belirsizlik = new ArrayList();

                    ArrayList arrayARFP_T7_Frekans = new ArrayList();
                    ArrayList arrayARFP_T7_Cıkıs_Gücü = new ArrayList();
                    ArrayList arrayARFP_T7_OlculenGuc = new ArrayList();
                    ArrayList arrayARFP_T7_AltSınır = new ArrayList();
                    ArrayList arrayARFP_T7_Sapma = new ArrayList();
                    ArrayList arrayARFP_T7_ÜstSınır = new ArrayList();
                    ArrayList arrayARFP_T7_Belirsizlik = new ArrayList();
                    ArrayList arrayARFP_T7_Belirsizlik2 = new ArrayList();

                    ArrayList arrayARFP_T8_Frekans = new ArrayList();
                    ArrayList arrayARFP_T8_Cıkıs_Gücü = new ArrayList();
                    ArrayList arrayARFP_T8_OlculenDeger = new ArrayList();
                    ArrayList arrayARFP_T8_AltSınır = new ArrayList();
                    ArrayList arrayARFP_T8_Fark = new ArrayList();
                    ArrayList arrayARFP_T8_ÜstSınır = new ArrayList();
                    ArrayList arrayARFP_T8_Belirsizlik = new ArrayList();
                    ArrayList arrayARFP_T8_Belirsizlik2 = new ArrayList();

                    ArrayList arrayARFP_T9_T10_T11_frekans = new ArrayList();
                    ArrayList arrayARFP_T9_SWR_Seviye = new ArrayList();
                    ArrayList arrayARFP_T9_SWR_OlculenDeger = new ArrayList();
                    ArrayList arrayARFP_T9_SWR_MaksimumDeger = new ArrayList();
                    ArrayList arrayARFP_T9_SWR_Belirsizlik = new ArrayList();

                    ArrayList arrayARFP_T10_SWR_Seviye = new ArrayList();
                    ArrayList arrayARFP_T10_SWR_OlculenDeger = new ArrayList();
                    ArrayList arrayARFP_T10_SWR_MaksimumDeger = new ArrayList();
                    ArrayList arrayARFP_T10_SWR_Belirsizlik = new ArrayList();

                    ArrayList arrayARFP_T11_SWR_Seviye = new ArrayList();
                    ArrayList arrayARFP_T11_SWR_OlculenDeger = new ArrayList();
                    ArrayList arrayARFP_T11_SWR_MaksimumDeger = new ArrayList();
                    ArrayList arrayARFP_T11_SWR_Belirsizlik = new ArrayList();
                    #endregion


                    if (boolList[0] == true)
                    {
                        FormatData(exd.ArrayS11Reel, exd.ArrayS11ReelUnc, arrays11reel, arrays11reelunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS11Complex, exd.ArrayS11ComplexUnc, arrays11complex, arrays11complexunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[1] == true)
                    {
                        FormatData(exd.ArrayS11Lin, exd.ArrayS11LinUnc, arrays11lin, arrays11linunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS11LinPhase, exd.ArrayS11LinPhaseUnc, arrays11linphase, arrays11linphaseunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[2] == true)
                    {
                        FormatData(exd.ArrayS11Log, exd.ArrayS11LogUnc, arrays11log, arrays11logunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS11LogPhase, exd.ArrayS11LogPhaseUnc, arrays11logphase, arrays11logphaseunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[3] == true)
                    {
                        FormatData(exd.ArrayS11SWR, exd.ArrayS11SWRUnc, arrays11swr, arrays11swrunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[4] == true)
                    {
                        FormatData(exd.ArrayS12Reel, exd.ArrayS12ReelUnc, arrays12reel, arrays12reelunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS12Complex, exd.ArrayS12ComplexUnc, arrays12complex, arrays12complexunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[5] == true)
                    {
                        FormatData(exd.ArrayS12Lin, exd.ArrayS12LinUnc, arrays12lin, arrays12linunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS12LinPhase, exd.ArrayS12LinPhaseUnc, arrays12linphase, arrays12linphaseunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[6] == true)
                    {
                        FormatData(exd.ArrayS12Log, exd.ArrayS12LogUnc, arrays12log, arrays12logunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS12LogPhase, exd.ArrayS12LogPhaseUnc, arrays12logphase, arrays12logphaseunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[7] == true)
                    {
                        FormatData(exd.ArrayS21Reel, exd.ArrayS21ReelUnc, arrays21reel, arrays21reelunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS21Complex, exd.ArrayS21ComplexUnc, arrays21complex, arrays21complexunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[8] == true)
                    {
                        FormatData(exd.ArrayS21Lin, exd.ArrayS21LinUnc, arrays21lin, arrays21linunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS21LinPhase, exd.ArrayS21LinPhaseUnc, arrays21linphase, arrays21linphaseunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[9] == true)
                    {
                        FormatData(exd.ArrayS21Log, exd.ArrayS21LogUnc, arrays21log, arrays21logunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS21LogPhase, exd.ArrayS21LogPhaseUnc, arrays21logphase, arrays21logphaseunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[10] == true)
                    {
                        FormatData(exd.ArrayS22Reel, exd.ArrayS22ReelUnc, arrays22reel, arrays22reelunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS22Complex, exd.ArrayS22ComplexUnc, arrays22complex, arrays22complexunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[11] == true)
                    {
                        FormatData(exd.ArrayS22Lin, exd.ArrayS22LinUnc, arrays22lin, arrays22linunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS22LinPhase, exd.ArrayS22LinPhaseUnc, arrays22linphase, arrays22linphaseunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[12] == true)
                    {
                        FormatData(exd.ArrayS22Log, exd.ArrayS22LogUnc, arrays22log, arrays22logunc, exd.ArrayFrekansSParam.Count);
                        FormatData(exd.ArrayS22LogPhase, exd.ArrayS22LogPhaseUnc, arrays22logphase, arrays22logphaseunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[13] == true)
                    {
                        FormatData(exd.ArrayS22SWR, exd.ArrayS22SWRUnc, arrays22swr, arrays22swrunc, exd.ArrayFrekansSParam.Count);
                    }
                    if (boolList[14] == true)
                    {
                        FormatData(exd.ArrayEffiencyEEEE, exd.ArrayEffiencyEEEEUnc, arraysEffiencyEEEE, arraysEffiencyEEEEunc, exd.ArrayFrekansEE.Count);
                    }
                    if (boolList[15] == true)
                    {
                        FormatData(exd.ArrayEffiencyEE_S11Imag, exd.ArrayEffiencyEE_S11ImagUnc, arraysEffiencyEE_S11Imag, arraysEffiencyEE_S11Imagunc, exd.ArrayFrekansEE.Count);
                        FormatData(exd.ArrayEffiencyEE_S11Reel, exd.ArrayEffiencyEE_S11ReelUnc, arraysEffiencyEE_S11Reel, arraysEffiencyEE_S11Reelunc, exd.ArrayFrekansEE.Count);
                        FormatData(exd.ArrayEffiencyRHO_EERho, exd.ArrayEffiencyRHO_EERhoUnc, arraysEffiencyRHO_EERho, arraysEffiencyRHO_EERhounc, exd.ArrayFrekansEE.Count);
                    }
                    if (boolList[16] == true)
                    {
                        FormatData(exd.ArrayEffiencyEE_CFEE_CF, exd.ArrayEffiencyEE_CFEE_CFUnc, arraysEffiencyEE_CFEE_CF, arraysEffiencyEE_CFEE_CFunc, exd.ArrayFrekansEE.Count);
                    }
                    if (boolList[17] == true)
                    {
                        FormatData(exd.ArrayCF_Cal_Factor, exd.ArrayCF_Cal_Factor_Unc, arrayCF_Cal_Factor, arrayCF_Cal_Factor_Unc, exd.ArrayFrekansCF.Count);
                    }
                    if (boolList[18] == true)
                    {
                        FormatData(exd.ArrayCF_Imaginer, exd.ArrayCF_Imaginer_Unc, arrayCF_Imaginer, arrayCF_Imaginer_Unc, exd.ArrayFrekansCF.Count);
                        FormatData(exd.ArrayCF_Reel, exd.ArrayCF_Reel_Unc, arrayCF_Reel, arrayCF_Reel_Unc, exd.ArrayFrekansCF.Count);
                        FormatData(exd.ArrayCF_ReflectionCof, exd.ArrayCF_ReflectionCof_Unc, arrayCF_ReflectionCof, arrayCF_ReflectionCof_Unc, exd.ArrayFrekansCF.Count);
                    }
                    if (boolList[19] == true)
                    {
                        FormatData(exd.ArrayCIS_Z_Position, exd.ArrayCIS_Z_Position_Unc, arrayCIS_Z_Position, arrayCIS_Z_Position_Unc, exd.ArrayFrekansCIS.Count);
                        FormatData(exd.ArrayCIS_ICOD, exd.ArrayCIS_ICOD_Unc, arrayCIS_ICOD, arrayCIS_ICOD_Unc, exd.ArrayFrekansCIS.Count);
                        FormatData(exd.ArrayCIS_OCID, exd.ArrayCIS_OCID_Unc, arrayCIS_OCID, arrayCIS_OCID_Unc, exd.ArrayFrekansCIS.Count);
                    }
                    if (boolList[20] == true)
                    {
                        FormatData(exd.ArrayNoiseENR, exd.ArrayNoiseENRUnc, arrayNoiseENR, arrayNoiseENRUnc, exd.ArrayFrekansNoise.Count);
                    }
                    if (boolList[21] == true)
                    {
                        FormatData(exd.ArrayNoiseDCONRCLin, exd.ArrayNoiseDCONRCLinUnc, arrayNoiseDCONRCLin, arrayNoiseDCONRCLinUnc, exd.ArrayFrekansNoise.Count);
                        FormatData(exd.ArrayNoiseDCONRCPhase, exd.ArrayNoiseDCONRCPhaseUnc, arrayNoiseDCONRCPhase, arrayNoiseDCONRCPhaseUnc, exd.ArrayFrekansNoise.Count);
                    }
                    if (boolList[22] == true)
                    {
                        FormatData(exd.ArrayNoiseDCOFFRCLin, exd.ArrayNoiseDCOFFRCLinUnc, arrayNoiseDCOFFRCLin, arrayNoiseDCOFFRCLinUnc, exd.ArrayFrekansNoise.Count);
                        FormatData(exd.ArrayNoiseDCOFFRCPhase, exd.ArrayNoiseDCOFFRCPhaseUnc, arrayNoiseDCOFFRCPhase, arrayNoiseDCOFFRCPhaseUnc, exd.ArrayFrekansNoise.Count);
                    }
                    if (boolList[23] == true)
                    {
                        FormatData(exd.ArrayRFD_T1_GostergeDegeri, exd.ArrayRFD_T1_Belirsizlik, arrayRFD_T1_GostergeDegeri, arrayRFD_T1_Belirsizlik, exd.ArrayRFD_T1_Frekans.Count);
                        FormatData(exd.ArrayRFD_T1_OlculenDeger, exd.ArrayRFD_T1_Belirsizlik, arrayRFD_T1_OlculenDeger, arrayRFD_T1_Belirsizlik2, exd.ArrayRFD_T1_Frekans.Count);
                        FormatData(exd.ArrayRFD_T1_OlculenFark, exd.ArrayRFD_T1_Belirsizlik, arrayRFD_T1_OlculenFark, arrayRFD_T1_Belirsizlik3, exd.ArrayRFD_T1_Frekans.Count);
                        arrayRFD_T1_AltSınır = exd.ArrayRFD_T1_AltSınır;
                        arrayRFD_T1_UstSınır = exd.ArrayRFD_T1_UstSınır;
                    }
                    if (boolList[24] == true)
                    {
                        FormatData(exd.ArrayRFD_T2_OlculenDeger, exd.ArrayRFD_T2_Belirsizlik, arrayRFD_T2_OlculenDeger, arrayRFD_T2_Belirsizlik, exd.ArrayRFD_T2_Frekans.Count);
                        FormatData(exd.ArrayRFD_T2_Nom_Guc_Lvl_fark, exd.ArrayRFD_T2_Belirsizlik, arrayRFD_T2_Nom_Guc_Lvl_fark, arrayRFD_T2_Belirsizlik2, exd.ArrayRFD_T2_Frekans.Count);
                        arrayRFD_T2_Nom_Guc_Lvl = exd.ArrayRFD_T2_Nom_Guc_Lvl;
                        arrayRFD_T2_AltSınır = exd.ArrayRFD_T2_AltSınır;
                        arrayRFD_T2_UstSınır = exd.ArrayRFD_T2_UstSınır;
                    }
                    if (boolList[25] == true)
                    {
                        FormatData(exd.ArrayRFD_T3_OlculenDeger, exd.ArrayRFD_T3_Belirsizlik, arrayRFD_T3_OlculenDeger, arrayRFD_T3_Belirsizlik, exd.ArrayRFD_T3_Frekans.Count);
                        FormatData(exd.ArrayRFD_T3_Fark, exd.ArrayRFD_T3_Belirsizlik, arrayRFD_T3_Fark, arrayRFD_T3_Belirsizlik2, exd.ArrayRFD_T3_Frekans.Count);
                        arrayRFD_T3_NominalGuc = exd.ArrayRFD_T3_NominalGuc;
                        arrayRFD_T3_AltSınır = exd.ArrayRFD_T3_AltSınır;
                        arrayRFD_T3_UstSınır = exd.ArrayRFD_T3_UstSınır;
                    }
                    if (boolList[26] == true)
                    {
                        FormatData(exd.ArrayRFD_T4_Fark, exd.ArrayRFD_T4_Belirsizlik, arrayRFD_T4_Fark, arrayRFD_T4_Belirsizlik, exd.ArrayRFD_T4_Frekans.Count);
                        arrayRFD_T4_Min_Guc_lvl = exd.ArrayRFD_T4_Min_Guc_lvl;
                        arrayRFD_T4_Max_Guc_lvl = exd.ArrayRFD_T4_Max_Guc_lvl;
                        arrayRFD_T4_AltSınır = exd.ArrayRFD_T4_AltSınır;
                        arrayRFD_T4_UstSınır = exd.ArrayRFD_T4_UstSınır;
                    }
                    if (boolList[27] == true)
                    {
                        FormatData(exd.ArrayRFG_T1_GirisGucu, exd.ArrayRFG_T1_Belirsizlik, arrayRFG_T1_GirisGucu, arrayRFG_T1_Belirsizlik, exd.ArrayRFG_T1_Frekans.Count);
             
                    }
                    if (boolList[28] == true)
                    {
                        FormatData(exd.ArrayRFG_T2_EnBuyukKazanc, exd.ArrayRFG_T2_EnKucukKazanc, arrayRFG_T2_EnBuyukKazanc, arrayRFG_T2_EnKucukKazanc, exd.ArrayRFG_T2_Flatness.Count);
                      
                    }
                    if (boolList[29] == true)
                    {
                        FormatData(exd.ArrayRFG_T3_Kazanc, exd.ArrayRFG_T3_Belirsizlik, arrayRFG_T3_Kazanc, arrayRFG_T3_Belirsizlik, exd.ArrayRFG_T3_Nom_Giris_Gucu.Count);
                     
                    }
                    if (boolList[30] == true)
                    {
                        FormatData(exd.ArrayRFG_T4_Kazanc, exd.ArrayRFG_T4_Belirsizlik, arrayRFG_T4_Kazanc, arrayRFG_T4_Belirsizlik, exd.ArrayRFG_T4_Nom_Giris_Gucu.Count);
                    
                    }
                    if (boolList[31] == true)
                    {
                        FormatData(exd.ArrayARFP_T1_Olculen_Güc, exd.ArrayARFP_T1_Belirsizlik, arrayARFP_T1_Olculen_Güc, arrayARFP_T1_Belirsizlik, exd.ArrayARFP_T1_Frekans.Count);
                        FormatData(exd.ArrayARFP_T1_Sapma, exd.ArrayARFP_T1_Belirsizlik, arrayARFP_T1_Sapma, arrayARFP_T1_Belirsizlik2, exd.ArrayARFP_T1_Frekans.Count);
                     arrayARFP_T1_Cıkıs_Gücü = exd.ArrayARFP_T1_Cıkıs_Gücü;
                        arrayARFP_T1_ÜstSınır = exd.ArrayARFP_T1_ÜstSınır;
                        arrayARFP_T1_AltSınır = exd.ArrayARFP_T1_AltSınır;
                    }
                    if (boolList[32] == true)
                    {
                        FormatData(exd.ArrayARFP_T2_OlculenDeger, exd.ArrayARFP_T2_Belirsizlik, arrayARFP_T2_OlculenDeger, arrayARFP_T2_Belirsizlik, exd.ArrayARFP_T2_Frekans.Count);
                        FormatData(exd.ArrayARFP_T2_Fark, exd.ArrayARFP_T2_Belirsizlik, arrayARFP_T2_Fark, arrayARFP_T2_Belirsizlik2, exd.ArrayARFP_T2_Frekans.Count);
                        arrayARFP_T2_Cıkıs_Gücü = exd.ArrayARFP_T2_Cıkıs_Gücü;
                        arrayARFP_T2_ÜstSınır = exd.ArrayARFP_T2_ÜstSınır;
                        arrayARFP_T2_AltSınır = exd.ArrayARFP_T2_AltSınır;
                 }
                    if (boolList[33] == true)
                    {
                        FormatData(exd.ArrayARFP_T3_OlculenZayıflatma, exd.ArrayARFP_T3_Belirsizlik, arrayARFP_T3_OlculenZayıflatma, arrayARFP_T3_Belirsizlik, exd.ArrayARFP_T3_Frekans.Count);
                        FormatData(exd.ArrayARFP_T3_Zayıflatma, exd.ArrayARFP_T3_Belirsizlik, arrayARFP_T3_Zayıflatma, arrayARFP_T3_Belirsizlik2, exd.ArrayARFP_T3_Frekans.Count);
                        arrayARFP_T3_Cıkıs_Gücü = exd.ArrayARFP_T3_Cıkıs_Gücü;
                        arrayARFP_T3_ÜstSınır = exd.ArrayARFP_T3_ÜstSınır;
                        arrayARFP_T3_AltSınır = exd.ArrayARFP_T3_AltSınır;
  }
                    if (boolList[34] == true)
                    {
                        FormatData(exd.ArrayARFP_T4_SWR_OlculenDeger, exd.ArrayARFP_T4_SWR_Belirsizlik, arrayARFP_T4_SWR_OlculenDeger, arrayARFP_T4_SWR_Belirsizlik, exd.ArrayARFP_T4_T5_T6_frekans.Count);
                        arrayARFP_T4_SWR_Seviye = exd.ArrayARFP_T4_SWR_Seviye;
                        arrayARFP_T4_SWR_MaksimumDeger = exd.ArrayARFP_T4_SWR_MaksimumDeger;
                    }
                    if (boolList[35] == true)
                    {
                        FormatData(exd.ArrayARFP_T5_SWR_OlculenDeger, exd.ArrayARFP_T5_SWR_Belirsizlik, arrayARFP_T5_SWR_OlculenDeger, arrayARFP_T5_SWR_Belirsizlik, exd.ArrayARFP_T4_T5_T6_frekans.Count);
                        arrayARFP_T5_SWR_Seviye = exd.ArrayARFP_T5_SWR_Seviye;
                        arrayARFP_T5_SWR_MaksimumDeger = exd.ArrayARFP_T5_SWR_MaksimumDeger;
                    }
                    if (boolList[36] == true)
                    {
                        FormatData(exd.ArrayARFP_T6_SWR_OlculenDeger, exd.ArrayARFP_T6_SWR_Belirsizlik, arrayARFP_T6_SWR_OlculenDeger, arrayARFP_T6_SWR_Belirsizlik, exd.ArrayARFP_T4_T5_T6_frekans.Count);
                        arrayARFP_T6_SWR_Seviye = exd.ArrayARFP_T6_SWR_Seviye;
                        arrayARFP_T6_SWR_MaksimumDeger = exd.ArrayARFP_T6_SWR_MaksimumDeger;
                    }
                    if (boolList[37] == true)
                    {
                        FormatData(exd.ArrayARFP_T7_OlculenGuc, exd.ArrayARFP_T7_Belirsizlik, arrayARFP_T7_OlculenGuc, arrayARFP_T7_Belirsizlik, exd.ArrayARFP_T7_Frekans.Count);
                        FormatData(exd.ArrayARFP_T7_Sapma, exd.ArrayARFP_T7_Belirsizlik, arrayARFP_T7_Sapma, arrayARFP_T7_Belirsizlik2, exd.ArrayARFP_T7_Frekans.Count);
                        arrayARFP_T7_Cıkıs_Gücü = exd.ArrayARFP_T7_Cıkıs_Gücü;
                        arrayARFP_T7_ÜstSınır = exd.ArrayARFP_T7_ÜstSınır;
                        arrayARFP_T7_AltSınır = exd.ArrayARFP_T7_AltSınır;
                    }
                    if (boolList[38] == true)
                    {
                        FormatData(exd.ArrayARFP_T8_OlculenDeger, exd.ArrayARFP_T8_Belirsizlik, arrayARFP_T8_OlculenDeger, arrayARFP_T8_Belirsizlik, exd.ArrayARFP_T8_Frekans.Count);
                        FormatData(exd.ArrayARFP_T8_Fark, exd.ArrayARFP_T8_Belirsizlik, arrayARFP_T8_Fark, arrayARFP_T8_Belirsizlik2, exd.ArrayARFP_T8_Frekans.Count);
                        arrayARFP_T8_Cıkıs_Gücü = exd.ArrayARFP_T8_Cıkıs_Gücü;
                        arrayARFP_T8_ÜstSınır = exd.ArrayARFP_T8_ÜstSınır;
                        arrayARFP_T8_AltSınır = exd.ArrayARFP_T8_AltSınır;
                    }
                    if (boolList[39] == true)
                    {
                        FormatData(exd.ArrayARFP_T9_SWR_OlculenDeger, exd.ArrayARFP_T9_SWR_Belirsizlik, arrayARFP_T9_SWR_OlculenDeger, arrayARFP_T9_SWR_Belirsizlik, exd.ArrayARFP_T9_T10_T11_frekans.Count);
                        arrayARFP_T9_SWR_Seviye = exd.ArrayARFP_T9_SWR_Seviye;
                        arrayARFP_T9_SWR_MaksimumDeger = exd.ArrayARFP_T9_SWR_MaksimumDeger;
                    }
                    if (boolList[40] == true)
                    {
                        FormatData(exd.ArrayARFP_T10_SWR_OlculenDeger, exd.ArrayARFP_T10_SWR_Belirsizlik, arrayARFP_T10_SWR_OlculenDeger, arrayARFP_T10_SWR_Belirsizlik, exd.ArrayARFP_T9_T10_T11_frekans.Count);
                        arrayARFP_T10_SWR_Seviye = exd.ArrayARFP_T10_SWR_Seviye;
                        arrayARFP_T10_SWR_MaksimumDeger = exd.ArrayARFP_T10_SWR_MaksimumDeger;
                    }
                    if (boolList[41] == true)
                    {
                        FormatData(exd.ArrayARFP_T11_SWR_OlculenDeger, exd.ArrayARFP_T11_SWR_Belirsizlik, arrayARFP_T11_SWR_OlculenDeger, arrayARFP_T11_SWR_Belirsizlik, exd.ArrayARFP_T9_T10_T11_frekans.Count);
                        arrayARFP_T11_SWR_MaksimumDeger = exd.ArrayARFP_T11_SWR_MaksimumDeger;
                        arrayARFP_T11_SWR_Seviye = exd.ArrayARFP_T11_SWR_Seviye;
                    }
                    

                    #endregion
                    #region Formatlanmış Verilerin Tablolara Dönüştürülmesi
                    if (boolList[0] == true)
                    {
                        Table table1 = ctb.CreateReelImg(exd.ArrayFrekansSParam, arrays11reel, arrays11reelunc, arrays11complex, arrays11complexunc);
                        tables.Add(table1);
                        this.headers.Add("Reel and Imaginary Components for S11\n\n");

                    }
                    if (boolList[1] == true)
                    {
                        Table table2 = ctb.CreateLinPhase(exd.ArrayFrekansSParam, arrays11lin, arrays11linunc, arrays11linphase, arrays11linphaseunc);
                        tables.Add(table2);
                        this.headers.Add("Linear Magnitude and Phase Components for S11\n\n");
                    }
                    if (boolList[2] == true)
                    {
                        Table table3 = ctb.CreateLogPhase(exd.ArrayFrekansSParam, arrays11log, arrays11logunc, arrays11logphase, arrays11logphaseunc);
                        tables.Add(table3);
                        this.headers.Add("Logarithmic Magnitude and Phase Components for S11\n\n");
                    }
                    if (boolList[3] == true)
                    {
                        Table table4 = ctb.CreateSWR(exd.ArrayFrekansSParam, arrays11swr, arrays11swrunc);
                        tables.Add(table4);
                        this.headers.Add("SWR Components for S11\n");
                    }
                    if (boolList[4] == true)
                    {
                        Table table5 = ctb.CreateReelImg(exd.ArrayFrekansSParam, arrays12reel, arrays12reelunc, arrays12complex, arrays12complexunc);
                        tables.Add(table5);
                        this.headers.Add("Reel and Imaginary Components for S12\n");
                    }
                    if (boolList[5] == true)
                    {
                        Table table6 = ctb.CreateLinPhase(exd.ArrayFrekansSParam, arrays12lin, arrays12linunc, arrays12linphase, arrays12linphaseunc);
                        tables.Add(table6);
                        this.headers.Add("Linear Magnitude and Phase Components for S12\n\n");
                    }
                    if (boolList[6] == true)
                    {
                        Table table7 = ctb.CreateLogPhase(exd.ArrayFrekansSParam, arrays12log, arrays12logunc, arrays12logphase, arrays12logphaseunc);
                        tables.Add(table7);
                        this.headers.Add("Logarithmic Magnitude and Phase Components for S12\n\n");
                    }
                    if (boolList[7] == true)
                    {
                        Table table8 = ctb.CreateReelImg(exd.ArrayFrekansSParam, arrays21reel, arrays21reelunc, arrays21complex, arrays21complexunc);
                        tables.Add(table8);
                        this.headers.Add("Reel and Imaginary Components for S21\n");
                    }
                    if (boolList[8] == true)
                    {
                        Table table9 = ctb.CreateLinPhase(exd.ArrayFrekansSParam, arrays21lin, arrays21linunc, arrays21linphase, arrays21linphaseunc);
                        tables.Add(table9);
                        this.headers.Add("Linear Magnitude and Phase Components for S21\n\n");
                    }
                    if (boolList[9] == true)
                    {
                        Table table10 = ctb.CreateLogPhase(exd.ArrayFrekansSParam, arrays21log, arrays21logunc, arrays21logphase, arrays21logphaseunc);
                        tables.Add(table10);
                        this.headers.Add("Logarithmic Magnitude and Phase Components for S21\n\n");
                    }
                    if (boolList[10] == true)
                    {
                        Table table11 = ctb.CreateReelImg(exd.ArrayFrekansSParam, arrays22reel, arrays22reelunc, arrays22complex, arrays22complexunc);
                        tables.Add(table11);
                        this.headers.Add("Reel and Imaginary Components for S22\n");
                    }
                    if (boolList[11] == true)
                    {
                        Table table12 = ctb.CreateLinPhase(exd.ArrayFrekansSParam, arrays22lin, arrays22linunc, arrays22linphase, arrays22linphaseunc);
                        tables.Add(table12);
                        this.headers.Add("Linear Magnitude and Phase Components for S22\n\n");
                    }
                    if (boolList[12] == true)
                    {
                        Table table13 = ctb.CreateLogPhase(exd.ArrayFrekansSParam, arrays22log, arrays22logunc, arrays22logphase, arrays22logphaseunc);
                        tables.Add(table13);
                        this.headers.Add("Logarithmic Magnitude and Phase Components for S22\n\n");
                    }
                    if (boolList[13] == true)
                    {
                        Table table14 = ctb.CreateSWR(exd.ArrayFrekansSParam, arrays22swr, arrays22swrunc);
                        tables.Add(table14);
                        this.headers.Add("SWR Components for S22\n");
                    }
                    if (boolList[14] == true)
                    {
                        Table table15 = ctb.CreateForTwoRow(exd.ArrayFrekansEE, arraysEffiencyEEEE, arraysEffiencyEEEEunc, "Ghz","EE 1", "EE 1 ING","EE 2","EE 2 ING");
                        tables.Add(table15);
                        this.headers.Add("EE 1\n");
                    }
                    if (boolList[15] == true)
                    {
                        Table table16 = ctb.CreateForSixRow(exd.ArrayFrekansEE, arraysEffiencyEE_S11Reel, arraysEffiencyEE_S11Reelunc, arraysEffiencyEE_S11Imag, arraysEffiencyEE_S11Imagunc, arraysEffiencyRHO_EERho, arraysEffiencyRHO_EERhounc);
                        tables.Add(table16);
                        this.headers.Add("EE 2\n");
                    }
                    if (boolList[16] == true)
                    {
                        Table table17 = ctb.CreateForTwoRow(exd.ArrayFrekansEE, arraysEffiencyEE_CFEE_CF, arraysEffiencyEE_CFEE_CFunc, "Ghz", "Kalibrasyon Fakötür", "Calibration Factor Unc,","Kalibrasyon Fakötür 2", "Calibration Factor Unc 2");
                        tables.Add(table17);
                        this.headers.Add("EE 3\n");
                    }
                    if (boolList[17] == true)
                    {
                        Table table18 = ctb.CreateForTwoRow(exd.ArrayFrekansCF, arrayCF_Cal_Factor, arrayCF_Cal_Factor_Unc, "Ghz","CF Kalibrasyon Faktörü", "CF Calibration Factor","CF Kalibrasyon Faktörü 2", "CF Calibration Factor 2");
                        tables.Add(table18);
                        this.headers.Add("CF 1 \n");
                    }
                    if (boolList[18] == true)
                    {
                        Table table19 = ctb.CreateForSixRow(exd.ArrayFrekansCF, arrayCF_Reel, arrayCF_Reel_Unc, arrayCF_Imaginer, arrayCF_Imaginer_Unc, arrayCF_ReflectionCof, arrayCF_ReflectionCof_Unc);
                        tables.Add(table19);
                        this.headers.Add("CF 2 \n");
                    }
                    if (boolList[19] == true)
                    {
                        Table table20 = ctb.CreateForSixRow(exd.ArrayFrekansCIS, arrayCIS_Z_Position, arrayCIS_Z_Position_Unc, arrayCIS_ICOD, arrayCIS_ICOD_Unc, arrayCIS_OCID, arrayCIS_OCID_Unc);
                        tables.Add(table20);
                        this.headers.Add("CF 3 \n");
                    }
                    if (boolList[20] == true)
                    {
                        Table table21 = ctb.CreateForTwoRow(exd.ArrayFrekansNoise, arrayNoiseENR, arrayNoiseENRUnc, "Ghz","Noise Enr", "Noise Enr Eng", "Nois2e Enr", "Nois2 Enr Eng");
                        tables.Add(table21);
                        this.headers.Add("Noise 1 \n");
                    }
                    if (boolList[21] == true)
                    {
                        Table table22 = ctb.CreateForFiveRow(exd.ArrayFrekansNoise, arrayNoiseDCONRCLinUnc, exd.ArrayNoiseDCONUpLimit, arrayNoiseDCONRCLinUnc, arrayNoiseDCONRCPhase, arrayNoiseDCONRCPhaseUnc);
                        tables.Add(table22);
                        this.headers.Add("Noise 2 \n");
                    }
                    if (boolList[22] == true)
                    {
                        Table table23 = ctb.CreateForFiveRow(exd.ArrayFrekansNoise, arrayNoiseDCOFFRCLinUnc, exd.ArrayDCOFFUpLimit, arrayNoiseDCOFFRCLinUnc, arrayNoiseDCOFFRCPhase, arrayNoiseDCOFFRCPhaseUnc);
                        tables.Add(table23);
                        this.headers.Add("Noise 3 \n");
                    }
                    if (boolList[23] == true)
                    {
                        Table table24 = ctb.CreateForSixRow(exd.ArrayRFD_T1_Frekans, arrayRFD_T1_GostergeDegeri, arrayRFD_T1_AltSınır, arrayRFD_T1_UstSınır, arrayRFD_T1_OlculenDeger, arrayRFD_T1_OlculenFark, arrayRFD_T1_Belirsizlik);
                        tables.Add(table24);
                        this.headers.Add("RF Diff Table 1 \n");
                    }
                    if (boolList[24] == true)
                    {
                        Table table25 = ctb.CreateForSixRow(exd.ArrayRFD_T2_Frekans, arrayRFD_T2_Nom_Guc_Lvl, arrayRFD_T2_OlculenDeger, arrayRFD_T2_AltSınır, arrayRFD_T2_Nom_Guc_Lvl_fark,arrayRFD_T2_UstSınır, arrayRFD_T2_Belirsizlik);
                        tables.Add(table25);
                        this.headers.Add("RF Diff Table 2 \n");
                    }
                    if (boolList[25] == true)
                    {
                        Table table26 = ctb.CreateForSixRow(exd.ArrayRFD_T3_Frekans, arrayRFD_T3_NominalGuc, arrayRFD_T3_AltSınır, arrayRFD_T3_OlculenDeger, arrayRFD_T3_UstSınır, arrayRFD_T3_Fark, arrayRFD_T3_Belirsizlik);
                        tables.Add(table26);
                        this.headers.Add("RF Diff Table 3 \n");
                    }
                    if (boolList[26] == true)
                    {
                        Table table27 = ctb.CreateForSixRow(exd.ArrayRFD_T4_Frekans, arrayRFD_T4_Min_Guc_lvl, arrayRFD_T4_Max_Guc_lvl, arrayRFD_T4_AltSınır, arrayRFD_T4_Fark, arrayRFD_T4_UstSınır,  arrayRFD_T4_Belirsizlik);
                        tables.Add(table27);
                        this.headers.Add("RF Diff Table 4 \n");
                    }
                    if (boolList[27] == true)
                    {
                        Table table28 = ctb.CreateSWR(exd.ArrayRFG_T1_Frekans, arrayRFG_T1_GirisGucu, arrayRFG_T1_Belirsizlik);
                        tables.Add(table28);
                        this.headers.Add("RF Gain Table 1 \n");
                    }
                    if (boolList[28] == true)
                    {
                        Table table29 = ctb.CreateSWR(exd.ArrayRFG_T2_EnBuyukKazanc, arrayRFG_T2_EnKucukKazanc, exd.ArrayRFG_T2_Flatness);
                        tables.Add(table29);
                        this.headers.Add("RF Gain Table 2 \n");
                    }
                    if (boolList[29] == true)
                    {
                        Table table30 = ctb.CreateSWR(exd.ArrayRFG_T3_Nom_Giris_Gucu, arrayRFG_T3_Kazanc, arrayRFG_T3_Belirsizlik);
                        tables.Add(table30);
                        this.headers.Add("RF Gain Table 3 \n");
                    }
                    if (boolList[30] == true)
                    {
                        Table table31 = ctb.CreateSWR(exd.ArrayRFG_T4_Nom_Giris_Gucu, arrayRFG_T4_Kazanc, arrayRFG_T4_Belirsizlik);
                        tables.Add(table31);
                        this.headers.Add("RF Gain Table 4 \n");
                    }
                    if (boolList[31] == true)
                    {
                        Table table = ctb.CreateForSixRow(exd.ArrayARFP_T1_Frekans, arrayARFP_T1_Cıkıs_Gücü, arrayARFP_T1_Olculen_Güc,arrayARFP_T1_AltSınır ,arrayARFP_T1_Sapma, arrayARFP_T1_ÜstSınır, arrayARFP_T1_Belirsizlik);
                        tables.Add(table);
                        this.headers.Add("Absolude 1 \n");
                    }
                    if (boolList[32] == true)
                    {
                        Table table = ctb.CreateForSixRow(exd.ArrayARFP_T2_Frekans, arrayARFP_T2_Cıkıs_Gücü, arrayARFP_T2_OlculenDeger, arrayARFP_T2_AltSınır, arrayARFP_T2_Fark, arrayARFP_T2_ÜstSınır, arrayARFP_T2_Belirsizlik);
                        tables.Add(table);
                        this.headers.Add("Absolude 2 \n");
                    }
                    if (boolList[33] == true)
                    {
                        Table table = ctb.CreateForSixRow(exd.ArrayARFP_T3_Frekans, arrayARFP_T3_Cıkıs_Gücü, arrayARFP_T3_OlculenZayıflatma, arrayARFP_T3_AltSınır, arrayARFP_T3_Zayıflatma, arrayARFP_T3_ÜstSınır, arrayARFP_T3_Belirsizlik);
                        tables.Add(table);
                        this.headers.Add("Absolude 3 \n");
                    }
                    if (boolList[34] == true)
                    {
                        Table table = ctb.CreateForFourRow(exd.ArrayARFP_T4_T5_T6_frekans, arrayARFP_T4_SWR_Seviye, arrayARFP_T4_SWR_OlculenDeger, arrayARFP_T4_SWR_MaksimumDeger, arrayARFP_T4_SWR_Belirsizlik, "Absolude 4", "Absolude 4", "Absolude 4", "Absolude 4");
                        tables.Add(table);
                        this.headers.Add("Absolude 4 \n");
                    }
                    if (boolList[35] == true)
                    {
                        Table table = ctb.CreateForFourRow(exd.ArrayARFP_T4_T5_T6_frekans, arrayARFP_T5_SWR_Seviye, arrayARFP_T5_SWR_OlculenDeger, arrayARFP_T5_SWR_MaksimumDeger, arrayARFP_T5_SWR_Belirsizlik, "Absolude 5", "Absolude 5", "Absolude 5", "Absolude 5");
                        tables.Add(table);
                        this.headers.Add("Absolude 5 \n");
                    }
                    if (boolList[36] == true)
                    {
                        Table table = ctb.CreateForFourRow(exd.ArrayARFP_T4_T5_T6_frekans, arrayARFP_T6_SWR_Seviye, arrayARFP_T6_SWR_OlculenDeger, arrayARFP_T6_SWR_MaksimumDeger, arrayARFP_T6_SWR_Belirsizlik, "Absolude 6", "Absolude 6", "Absolude 6", "Absolude 6");
                        tables.Add(table);
                        this.headers.Add("Absolude 6 \n");
                    }
                    if (boolList[37] == true)
                    {
                        Table table = ctb.CreateForSixRow(exd.ArrayARFP_T7_Frekans, arrayARFP_T7_Cıkıs_Gücü, arrayARFP_T7_OlculenGuc, arrayARFP_T7_AltSınır, arrayARFP_T7_Sapma, arrayARFP_T7_ÜstSınır, arrayARFP_T7_Belirsizlik);
                        tables.Add(table);
                        this.headers.Add("Absolude 7 \n");
                    }
                    if (boolList[38] == true)
                    {
                        Table table = ctb.CreateForSixRow(exd.ArrayARFP_T8_Frekans, arrayARFP_T8_Cıkıs_Gücü, arrayARFP_T8_OlculenDeger, arrayARFP_T8_AltSınır, arrayARFP_T8_Fark, arrayARFP_T8_ÜstSınır, arrayARFP_T8_Belirsizlik);
                        tables.Add(table);
                        this.headers.Add("Absolude 8 \n");
                    }
                    if (boolList[39] == true)
                    {
                        Table table = ctb.CreateForFourRow(exd.ArrayARFP_T9_T10_T11_frekans, arrayARFP_T9_SWR_Seviye, arrayARFP_T9_SWR_OlculenDeger, arrayARFP_T9_SWR_MaksimumDeger, arrayARFP_T9_SWR_Belirsizlik, "Absolude 9", "Absolude 9", "Absolude 9", "Absolude 9");
                        tables.Add(table);
                        this.headers.Add("Absolude 9 \n");
                    }
                    if (boolList[40] == true)
                    {
                        Table table = ctb.CreateForFourRow(exd.ArrayARFP_T9_T10_T11_frekans, arrayARFP_T10_SWR_Seviye, arrayARFP_T10_SWR_OlculenDeger, arrayARFP_T10_SWR_MaksimumDeger, arrayARFP_T10_SWR_Belirsizlik, "Absolude 10", "Absolude 10", "Absolude 10", "Absolude 10");
                        tables.Add(table);
                        this.headers.Add("Absolude 10 \n");
                    }
                    if (boolList[41] == true)
                    {
                        Table table = ctb.CreateForFourRow(exd.ArrayARFP_T9_T10_T11_frekans, arrayARFP_T11_SWR_Seviye, arrayARFP_T11_SWR_OlculenDeger, arrayARFP_T11_SWR_MaksimumDeger, arrayARFP_T11_SWR_Belirsizlik, "Absolude 11", "Absolude 11", "Absolude 11", "Absolude 11");
                        tables.Add(table);
                        this.headers.Add("Absolude 11 \n");
                    }

                    exd.ClearData();
                    #endregion
                }
                ctb.ResultPages(tables);
            }
            catch (Exception ex)
            {
                // Hata durumunda hata mesajını yazdırın
                Console.WriteLine("Hata oluştu: " + ex.Message);
            }
        }
        public static List<bool> SelectFilledColumns(XElement resultElement)
        {
            List<bool> boolList = new List<bool>(new bool[42]);

            XNamespace dcc = "https://ptb.de/dcc";


            foreach (var quantityElement in resultElement.Descendants(dcc + "quantity"))
            {
                // refType değerinin boş olup olmadığını kontrol et
                if (quantityElement.Attribute("refType") != null)
                {
                    // Eğer refType="s_parameters11Reel" ise boolList'in 0. indeksi true olsun
                    if (quantityElement.Attribute("refType").Value == "s_parameters11Reel")
                    {
                        boolList[0] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters11Lin")
                    {
                        boolList[1] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters11Log")
                    {
                        boolList[2] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters11swr")
                    {
                        boolList[3] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters12Reel")
                    {
                        boolList[4] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters12Lin")
                    {
                        boolList[5] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters12Log")
                    {
                        boolList[6] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters21Reel")
                    {
                        boolList[7] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters21Lin")
                    {
                        boolList[8] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters21Log")
                    {
                        boolList[9] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters22Reel")
                    {
                        boolList[10] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters22Lin")
                    {
                        boolList[11] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters22Log")
                    {
                        boolList[12] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "s_parameters22swr")
                    {
                        boolList[13] = true;
                    }

                    if (quantityElement.Attribute("refType").Value == "Effective Effiency EE-EE")
                    {
                        boolList[14] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "Effective Effiency EE-Reel")
                    {
                        boolList[15] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "Effective Effiency EE-Cal_Factor")
                    {
                        boolList[16] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "Calibration Factor CF-Cal_Factor")
                    {
                        boolList[17] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "Calibration Factor CF-Reel")
                    {
                        boolList[18] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "Calculable Impedance Standard CIS-Z-Position")
                    {
                        boolList[19] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "Noise_ENR")
                    {
                        boolList[20] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "Noise_DC_ON_Lin")
                    {
                        boolList[21] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "Noise_DC_OFF_Lin")
                    {
                        boolList[22] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "RF_Diff_t1")
                    {
                        boolList[23] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "RF_Diff_t2")
                    {
                        boolList[24] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "RF_Diff_t3")
                    {
                        boolList[25] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "RF_Diff_t4")
                    {
                        boolList[26] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "RFG_Unc1")
                    {
                        boolList[27] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "RFG_Flatness")
                    {
                        boolList[28] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "RFG_Unc2")
                    {
                        boolList[29] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "RFG_Unc3")
                    {
                        boolList[30] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "frequency_ARFP_t1")
                    {
                        boolList[31] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "frequency_ARFP_t2")
                    {
                        boolList[32] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "frequency_ARFP_t3")
                    {
                        boolList[33] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "frequency_ARFP_t4")
                    {
                        boolList[34] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "frequency_ARFP_t5")
                    {
                        boolList[35] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "frequency_ARFP_t6")
                    {
                        boolList[36] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "frequency_ARFP_t7")
                    {
                        boolList[37] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "frequency_ARFP_t8")
                    {
                        boolList[38] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "frequency_ARFP_t9")
                    {
                        boolList[39] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "frequency_ARFP_t10")
                    {
                        boolList[40] = true;
                    }
                    if (quantityElement.Attribute("refType").Value == "frequency_ARFP_t11")
                    {
                        boolList[41] = true;
                    }

                }
            }

            return boolList;
        }
        static void FormatData(ArrayList ArrayMeasurent, ArrayList ArrayUncertainty, ArrayList ArrayMsrt, ArrayList ArrayUnc, int counter)
        {
            ExcelData exd = new ExcelData();
            NumberFormatter formatter = new NumberFormatter();
            CalculateEntity calculateEntity = new CalculateEntity();
            for (int i = 0; i < counter; i++)
            {
                calculateEntity.measurent = Convert.ToDecimal(ArrayMeasurent[i]);
                calculateEntity.uncertainty = Convert.ToDecimal(ArrayUncertainty[i]);
                CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                ArrayMsrt.Add(formattedEntity.measurent);
                ArrayUnc.Add(formattedEntity.uncertainty);
            }
        }

    }
}
