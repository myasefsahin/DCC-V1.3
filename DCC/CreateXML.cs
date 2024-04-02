using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections;

using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Drawing;


namespace DCC
{
    class CreateXML
    {
        private string dcc = "https://ptb.de/dcc";
        private string si = "https://ptb.de/si";
        public List<bool> dataList = new List<bool>();
        bool CIS_bool;
        public List<string> headers = new List<string>();
        XML_Arrays dataXml = new XML_Arrays();




        private string orderNo = "/dcc:digitalCalibrationCertificate/dcc:administrativeData/dcc:coreData/dcc:identifications/dcc:identification[@id='orderno']/dcc:value";
        private string itemName = "/dcc:digitalCalibrationCertificate/dcc:administrativeData/dcc:items/dcc:item/dcc:name[@id='itemname']/dcc:content";
        private string itemSerialNumber = "/dcc:digitalCalibrationCertificate/dcc:administrativeData/dcc:items/dcc:item/dcc:identifications/dcc:identification[@id='serialnumber']/dcc:value";

        #region S-PARAMETRE XML İŞLEMLERİ

        #region Administrative Data
        public XmlDocument AddAdministrativeData(XmlDocument xml, XML_Arrays dataXml)
        {
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            XmlNode orderNoNode = xml.SelectSingleNode(orderNo, nsmgr);
            XmlNode itemNameNode = xml.SelectSingleNode(itemName, nsmgr);
            XmlNode itemSerialNoNode = xml.SelectSingleNode(itemSerialNumber, nsmgr);

            string orderNoData = dataXml.XmlOrderNumber;
            string itemNameData = dataXml.XmlDeviceName;
            string itemSerialNoData = dataXml.XmlSerialNumber;

            orderNoNode.InnerText = orderNoData;
            itemNameNode.InnerText = itemNameData;
            itemSerialNoNode.InnerText = itemSerialNoData;

            return xml;
        }
        #endregion

        #region S_Parameter_Result
        // S Parametre için results elementi altına result elemeti oluşturma.
        public XmlDocument AddSParameterResult(XmlDocument xml, string str, XML_Arrays dataXml, List<bool> control)
        {
            //Result Namespace oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", dcc);
            nsmgr.AddNamespace("si", si);

            this.dataList = control;

            XmlNode sResults = xml.SelectSingleNode("/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results", nsmgr);

            // Elementlerin oluşturulması
            XmlElement result = xml.CreateElement("dcc", "result", dcc);
            result.SetAttribute("id", "s_parameter" + str + "_dB");
            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement data = xml.CreateElement("dcc", "data", dcc);
            XmlElement list = xml.CreateElement("dcc", "list", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "s_parameter_" + str + "_dB";

            //Dataları içeren elementler oluşturulup dcc:List elementine eklenir.
            list.AppendChild(addSParameterFrekans(xml, dataXml));

            if (control[0])
            {
                List<XmlElement> xmlList = AddReelImg(xml, dataXml.XmlArrayS11Reel, dataXml.XmlArrayS11ReelUnc, dataXml.XmlArrayS11Complex, dataXml.XmlArrayS11ComplexUnc, "s11");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[1])
            {
                List<XmlElement> xmlList = AddLinPhase(xml, dataXml.XmlArrayS11Lin, dataXml.XmlArrayS11LinUnc, dataXml.XmlArrayS11LinPhase, dataXml.XmlArrayS11LinPhaseUnc, "s11");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[2])
            {
                List<XmlElement> xmlList = AddLogPhase(xml, dataXml.XmlArrayS11Log, dataXml.XmlArrayS11LogUnc, dataXml.XmlArrayS11LogPhase, dataXml.XmlArrayS11LogPhaseUnc, "s11");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[3])
            {
                List<XmlElement> xmlList = AddSwr(xml, dataXml.XmlArrayS11SWR, dataXml.XmlArrayS11SWRUnc, "s11");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[4])
            {
                List<XmlElement> xmlList = AddReelImg(xml, dataXml.XmlArrayS12Reel, dataXml.XmlArrayS12ReelUnc, dataXml.XmlArrayS12Complex, dataXml.XmlArrayS12ComplexUnc, "s12");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[5])
            {
                List<XmlElement> xmlList = AddLinPhase(xml, dataXml.XmlArrayS12Lin, dataXml.XmlArrayS12LinUnc, dataXml.XmlArrayS12LinPhase, dataXml.XmlArrayS12LinPhaseUnc, "s12");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[6])
            {
                List<XmlElement> xmlList = AddLogPhase(xml, dataXml.XmlArrayS12Log, dataXml.XmlArrayS12LogUnc, dataXml.XmlArrayS12LogPhase, dataXml.XmlArrayS12LogPhaseUnc, "s12");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[7])
            {
                List<XmlElement> xmlList = AddReelImg(xml, dataXml.XmlArrayS21Reel, dataXml.XmlArrayS21ReelUnc, dataXml.XmlArrayS21Complex, dataXml.XmlArrayS21ComplexUnc, "s21");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[8])
            {
                List<XmlElement> xmlList = AddLinPhase(xml, dataXml.XmlArrayS21Lin, dataXml.XmlArrayS21LinUnc, dataXml.XmlArrayS21LinPhase, dataXml.XmlArrayS21LinPhaseUnc, "s21");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[9])
            {
                List<XmlElement> xmlList = AddLogPhase(xml, dataXml.XmlArrayS21Log, dataXml.XmlArrayS21LogUnc, dataXml.XmlArrayS21LogPhase, dataXml.XmlArrayS21LogPhaseUnc, "s21");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[10])
            {
                List<XmlElement> xmlList = AddReelImg(xml, dataXml.XmlArrayS22Reel, dataXml.XmlArrayS22ReelUnc, dataXml.XmlArrayS22Complex, dataXml.XmlArrayS22ComplexUnc, "s22");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[11])
            {
                List<XmlElement> xmlList = AddLinPhase(xml, dataXml.XmlArrayS22Lin, dataXml.XmlArrayS22LinUnc, dataXml.XmlArrayS22LinPhase, dataXml.XmlArrayS22LinPhaseUnc, "s22");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[12])
            {
                List<XmlElement> xmlList = AddLogPhase(xml, dataXml.XmlArrayS22Log, dataXml.XmlArrayS22LogUnc, dataXml.XmlArrayS22LogPhase, dataXml.XmlArrayS22LogPhaseUnc, "s22");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[13])
            {
                List<XmlElement> xmlList = AddSwr(xml, dataXml.XmlArrayS22SWR, dataXml.XmlArrayS22SWRUnc, "s22");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }


            //Son eklemeler yapılarak data geçirme tamamlanır.
            name.AppendChild(content);
            result.AppendChild(name);
            data.AppendChild(list);
            result.AppendChild(data);
            sResults.AppendChild(result);

            return xml;
        }
        #endregion

        #region S_Parameter_Frequency
        public XmlElement addSParameterFrekans(XmlDocument xml, XML_Arrays dataXml)
        {
            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            XmlElement frekansElement = xml.CreateElement("dcc", "quantity", dcc);
            frekansElement.SetAttribute("refType", "frequency_sp");

            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "Frequency";

            XmlElement hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement realList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement unit = xml.CreateElement("si", "unitXMLList", si);
            unit.InnerText = "\\dB";

            string frekansData = string.Join(" ", dataXml.XmlArrayFrekans.ToArray());
            value.InnerText = frekansData;

            name.AppendChild(content);
            realList.AppendChild(value);
            realList.AppendChild(unit);
            hibrid.AppendChild(realList);
            frekansElement.AppendChild(name);
            frekansElement.AppendChild(hibrid);

            return frekansElement;

        }
        #endregion

        #region S_Parameter
        public List<XmlElement> AddReelImg(XmlDocument xml, ArrayList reel, ArrayList reelUnc, ArrayList imag, ArrayList imagUnc, string sParameter)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // Reel Element Oluşturulması
            XmlElement reelElement = xml.CreateElement("dcc", "quantity", dcc);
            reelElement.SetAttribute("refType", "s_parameter" + sParameter + "Reel");

            XmlElement reelName = xml.CreateElement("dcc", "name", dcc);
            XmlElement reelContent = xml.CreateElement("dcc", "content", dcc);
            reelContent.InnerText = sParameter + " Reel";

            XmlElement reelHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement reelRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement reelValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement reelUnit = xml.CreateElement("si", "unitXMLList", si);
            reelUnit.InnerText = "\\dB";

            string reelData = string.Join(" ", reel.ToArray());
            reelValue.InnerText = reelData;

            reelName.AppendChild(reelContent);
            reelRealList.AppendChild(reelValue);
            reelRealList.AppendChild(reelUnit);
            reelHibrid.AppendChild(reelRealList);
            reelElement.AppendChild(reelName);
            reelElement.AppendChild(reelHibrid);

            xmlElements.Add(reelElement);

            // Reel Unc Element Oluşturulması
            XmlElement reelUncElement = xml.CreateElement("dcc", "quantity", dcc);
            reelUncElement.SetAttribute("refType", "s_parameter" + sParameter + "ReelUnc");

            XmlElement reelUncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement reelUncContent = xml.CreateElement("dcc", "content", dcc);
            reelUncContent.InnerText = sParameter + " ReelUnc";

            XmlElement reelUncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement reelUncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement reelUncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement reelUncUnit = xml.CreateElement("si", "unitXMLList", si);
            reelUncUnit.InnerText = "\\dB";

            string reelUncData = string.Join(" ", reelUnc.ToArray());
            reelUncValue.InnerText = reelUncData;

            reelUncName.AppendChild(reelUncContent);
            reelUncRealList.AppendChild(reelUncValue);
            reelUncRealList.AppendChild(reelUncUnit);
            reelUncHibrid.AppendChild(reelUncRealList);
            reelUncElement.AppendChild(reelUncName);
            reelUncElement.AppendChild(reelUncHibrid);

            xmlElements.Add(reelUncElement);

            // Imaginary Element Oluşturulması
            XmlElement imagElement = xml.CreateElement("dcc", "quantity", dcc);
            imagElement.SetAttribute("refType", "s_parameter" + sParameter + "Imag");

            XmlElement imagName = xml.CreateElement("dcc", "name", dcc);
            XmlElement imagContent = xml.CreateElement("dcc", "content", dcc);
            imagContent.InnerText = sParameter + " Imag";

            XmlElement imagHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement imagRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement imagValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement imagUnit = xml.CreateElement("si", "unitXMLList", si);
            imagUnit.InnerText = "\\dB";

            string imagData = string.Join(" ", imag.ToArray());
            imagValue.InnerText = imagData;

            imagName.AppendChild(imagContent);
            imagRealList.AppendChild(imagValue);
            imagRealList.AppendChild(imagUnit);
            imagHibrid.AppendChild(imagRealList);
            imagElement.AppendChild(imagName);
            imagElement.AppendChild(imagHibrid);

            xmlElements.Add(imagElement);

            // Imaginary Unc Element Oluşturulması
            XmlElement imagUncElement = xml.CreateElement("dcc", "quantity", dcc);
            imagUncElement.SetAttribute("refType", "s_parameter" + sParameter + "ImagUnc");

            XmlElement imagUncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement imagUncContent = xml.CreateElement("dcc", "content", dcc);
            imagUncContent.InnerText = sParameter + " ImagUnc";

            XmlElement imagUncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement imagUncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement imagUncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement imagUncUnit = xml.CreateElement("si", "unitXMLList", si);
            imagUncUnit.InnerText = "\\dB";

            string imagUncData = string.Join(" ", imagUnc.ToArray());
            imagUncValue.InnerText = imagUncData;

            imagUncName.AppendChild(imagUncContent);
            imagUncRealList.AppendChild(imagUncValue);
            imagUncRealList.AppendChild(imagUncUnit);
            imagUncHibrid.AppendChild(imagUncRealList);
            imagUncElement.AppendChild(imagUncName);
            imagUncElement.AppendChild(imagUncHibrid);

            xmlElements.Add(imagUncElement);

            return xmlElements;
        }

        public List<XmlElement> AddLinPhase(XmlDocument xml, ArrayList lin, ArrayList linUnc, ArrayList phase, ArrayList phaseUnc, string sParameter)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // Lin Element Oluşturulması
            XmlElement linElement = xml.CreateElement("dcc", "quantity", dcc);
            linElement.SetAttribute("refType", "s_parameter" + sParameter + "Lin");

            XmlElement linName = xml.CreateElement("dcc", "name", dcc);
            XmlElement linContent = xml.CreateElement("dcc", "content", dcc);
            linContent.InnerText = sParameter + " Lin";

            XmlElement linHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement linRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement linValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement linUnit = xml.CreateElement("si", "unitXMLList", si);
            linUnit.InnerText = "\\dB";

            string linData = string.Join(" ", lin.ToArray());
            linValue.InnerText = linData;

            linName.AppendChild(linContent);
            linRealList.AppendChild(linValue);
            linRealList.AppendChild(linUnit);
            linHibrid.AppendChild(linRealList);
            linElement.AppendChild(linName);
            linElement.AppendChild(linHibrid);

            xmlElements.Add(linElement);

            // Lin Unc Element Oluşturulması
            XmlElement linUncElement = xml.CreateElement("dcc", "quantity", dcc);
            linUncElement.SetAttribute("refType", "s_parameter" + sParameter + "LinUnc");

            XmlElement linUncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement linUncContent = xml.CreateElement("dcc", "content", dcc);
            linUncContent.InnerText = sParameter + " Lin Unc";

            XmlElement linUncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement linUncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement linUncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement linUncUnit = xml.CreateElement("si", "unitXMLList", si);
            linUncUnit.InnerText = "\\dB";

            string linUncData = string.Join(" ", linUnc.ToArray());
            linUncValue.InnerText = linUncData;

            linUncName.AppendChild(linUncContent);
            linUncRealList.AppendChild(linUncValue);
            linUncRealList.AppendChild(linUncUnit);
            linUncHibrid.AppendChild(linUncRealList);
            linUncElement.AppendChild(linUncName);
            linUncElement.AppendChild(linUncHibrid);

            xmlElements.Add(linUncElement);

            // Phase Element Oluşturulması
            XmlElement phaseElement = xml.CreateElement("dcc", "quantity", dcc);
            phaseElement.SetAttribute("refType", "s_parameter" + sParameter + "Phase");

            XmlElement phaseName = xml.CreateElement("dcc", "name", dcc);
            XmlElement phaseContent = xml.CreateElement("dcc", "content", dcc);
            phaseContent.InnerText = sParameter + " Phase";

            XmlElement phaseHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement phaseRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement phaseValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement phaseUnit = xml.CreateElement("si", "unitXMLList", si);
            phaseUnit.InnerText = "\\dB";

            string phaseData = string.Join(" ", phase.ToArray());
            phaseValue.InnerText = phaseData;

            phaseName.AppendChild(phaseContent);
            phaseRealList.AppendChild(phaseValue);
            phaseRealList.AppendChild(phaseUnit);
            phaseHibrid.AppendChild(phaseRealList);
            phaseElement.AppendChild(phaseName);
            phaseElement.AppendChild(phaseHibrid);

            xmlElements.Add(phaseElement);

            // phase Unc Element Oluşturulması
            XmlElement phaseUncElement = xml.CreateElement("dcc", "quantity", dcc);
            phaseUncElement.SetAttribute("refType", "s_parameter" + sParameter + "PhaseUnc");

            XmlElement phaseUncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement phaseUncContent = xml.CreateElement("dcc", "content", dcc);
            phaseUncContent.InnerText = sParameter + " PhaseUnc";

            XmlElement phaseUncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement phaseUncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement phaseUncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement phaseUncUnit = xml.CreateElement("si", "unitXMLList", si);
            phaseUncUnit.InnerText = "\\dB";

            string phaseUncData = string.Join(" ", phaseUnc.ToArray());
            phaseUncValue.InnerText = phaseUncData;

            phaseUncName.AppendChild(phaseUncContent);
            phaseUncRealList.AppendChild(phaseUncValue);
            phaseUncRealList.AppendChild(phaseUncUnit);
            phaseUncHibrid.AppendChild(phaseUncRealList);
            phaseUncElement.AppendChild(phaseUncName);
            phaseUncElement.AppendChild(phaseUncHibrid);

            xmlElements.Add(phaseUncElement);

            return xmlElements;
        }

        public List<XmlElement> AddLogPhase(XmlDocument xml, ArrayList log, ArrayList logUnc, ArrayList logPhase, ArrayList logPhaseUnc, string sParameter)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // Log Element Oluşturulması
            XmlElement logElement = xml.CreateElement("dcc", "quantity", dcc);
            logElement.SetAttribute("refType", "s_parameter" + sParameter + "Log");


            XmlElement logName = xml.CreateElement("dcc", "name", dcc);
            XmlElement logContent = xml.CreateElement("dcc", "content", dcc);
            logContent.InnerText = sParameter + " Log";

            XmlElement logHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement logRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement logValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement logUnit = xml.CreateElement("si", "unitXMLList", si);
            logUnit.InnerText = "\\dB";

            string logData = string.Join(" ", log.ToArray());
            logValue.InnerText = logData;

            logName.AppendChild(logContent);
            logRealList.AppendChild(logValue);
            logRealList.AppendChild(logUnit);
            logHibrid.AppendChild(logRealList);
            logElement.AppendChild(logName);
            logElement.AppendChild(logHibrid);

            xmlElements.Add(logElement);

            // Log Unc Element Oluşturulması
            XmlElement logUncElement = xml.CreateElement("dcc", "quantity", dcc);
            logUncElement.SetAttribute("refType", "s_parameter" + sParameter + "LogUnc");

            XmlElement logUncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement logUncContent = xml.CreateElement("dcc", "content", dcc);
            logUncContent.InnerText = sParameter + " Log Unc";

            XmlElement logUncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement logUncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement logUncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement logUncUnit = xml.CreateElement("si", "unitXMLList", si);
            logUncUnit.InnerText = "\\dB";

            string logUncData = string.Join(" ", logUnc.ToArray());
            logUncValue.InnerText = logUncData;

            logUncName.AppendChild(logUncContent);
            logUncRealList.AppendChild(logUncValue);
            logUncRealList.AppendChild(logUncUnit);
            logUncHibrid.AppendChild(logUncRealList);
            logUncElement.AppendChild(logUncName);
            logUncElement.AppendChild(logUncHibrid);

            xmlElements.Add(logUncElement);

            // Log Phase Element Oluşturulması
            XmlElement logPhaseElement = xml.CreateElement("dcc", "quantity", dcc);
            logPhaseElement.SetAttribute("refType", "s_parameter" + sParameter + "LogPhase");

            XmlElement logPhaseName = xml.CreateElement("dcc", "name", dcc);
            XmlElement logPhaseContent = xml.CreateElement("dcc", "content", dcc);
            logPhaseContent.InnerText = sParameter + " Log Phase";

            XmlElement logPhaseHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement logPhaseRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement logPhaseValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement logPhaseUnit = xml.CreateElement("si", "unitXMLList", si);
            logPhaseUnit.InnerText = "\\dB";

            string logPhaseData = string.Join(" ", logPhase.ToArray());
            logPhaseValue.InnerText = logPhaseData;

            logPhaseName.AppendChild(logPhaseContent);
            logPhaseRealList.AppendChild(logPhaseValue);
            logPhaseRealList.AppendChild(logPhaseUnit);
            logPhaseHibrid.AppendChild(logPhaseRealList);
            logPhaseElement.AppendChild(logPhaseName);
            logPhaseElement.AppendChild(logPhaseHibrid);

            xmlElements.Add(logPhaseElement);

            // LogPhase Unc Element Oluşturulması
            XmlElement logPhaseUncElement = xml.CreateElement("dcc", "quantity", dcc);
            logPhaseUncElement.SetAttribute("refType", "s_parameter" + sParameter + "LogPhaseUnc");

            XmlElement logPhaseUncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement logPhaseUncContent = xml.CreateElement("dcc", "content", dcc);
            logPhaseUncContent.InnerText = sParameter + " Log Phase Unc";

            XmlElement logPhaseUncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement logPhaseUncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement logPhaseUncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement logPhaseUncUnit = xml.CreateElement("si", "unitXMLList", si);
            logPhaseUncUnit.InnerText = "\\dB";

            string logPhaseUncData = string.Join(" ", logPhaseUnc.ToArray());
            logPhaseUncValue.InnerText = logPhaseUncData;

            logPhaseUncName.AppendChild(logPhaseUncContent);
            logPhaseUncRealList.AppendChild(logPhaseUncValue);
            logPhaseUncRealList.AppendChild(logPhaseUncUnit);
            logPhaseUncHibrid.AppendChild(logPhaseUncRealList);
            logPhaseUncElement.AppendChild(logPhaseUncName);
            logPhaseUncElement.AppendChild(logPhaseUncHibrid);

            xmlElements.Add(logPhaseUncElement);

            return xmlElements;
        }

        public List<XmlElement> AddSwr(XmlDocument xml, ArrayList swr, ArrayList swrUnc, string sParameter)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // SWR Element Oluşturulması
            XmlElement swrElement = xml.CreateElement("dcc", "quantity", dcc);
            swrElement.SetAttribute("refType", "s_parameter" + sParameter + "swr");

            XmlElement swrName = xml.CreateElement("dcc", "name", dcc);
            XmlElement swrContent = xml.CreateElement("dcc", "content", dcc);
            swrContent.InnerText = sParameter + " SWR";

            XmlElement swrHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement swrRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement swrValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement swrUnit = xml.CreateElement("si", "unitXMLList", si);
            swrUnit.InnerText = "\\dB";

            string swrData = string.Join(" ", swr.ToArray());
            swrValue.InnerText = swrData;

            swrName.AppendChild(swrContent);
            swrRealList.AppendChild(swrValue);
            swrRealList.AppendChild(swrUnit);
            swrHibrid.AppendChild(swrRealList);
            swrElement.AppendChild(swrName);
            swrElement.AppendChild(swrHibrid);

            xmlElements.Add(swrElement);

            // SWR Unc Element Oluşturulması
            XmlElement swrUncElement = xml.CreateElement("dcc", "quantity", dcc);
            swrUncElement.SetAttribute("refType", "s_parameter" + sParameter + "swrUnc");

            XmlElement swrUncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement swrUncContent = xml.CreateElement("dcc", "content", dcc);
            swrUncContent.InnerText = sParameter + " SWR Unc";

            XmlElement swrUncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement swrUncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement swrUncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement swrUncUnit = xml.CreateElement("si", "unitXMLList", si);
            swrUncUnit.InnerText = "\\dB";

            string swrUncData = string.Join(" ", swrUnc.ToArray());
            swrUncValue.InnerText = swrUncData;

            swrUncName.AppendChild(swrUncContent);
            swrUncRealList.AppendChild(swrUncValue);
            swrUncRealList.AppendChild(swrUncUnit);
            swrUncHibrid.AppendChild(swrUncRealList);
            swrUncElement.AppendChild(swrUncName);
            swrUncElement.AppendChild(swrUncHibrid);

            xmlElements.Add(swrUncElement);

            return xmlElements;
        }
        #endregion
        #endregion

        #region EE XML İŞLEMLERİ
        public XmlDocument Add_EE_Result(XmlDocument xml, string str, XML_Arrays dataXml, List<bool> control)
        {
            //Result Namespace oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", dcc);
            nsmgr.AddNamespace("si", si);

            this.dataList = control;

            XmlNode sResults = xml.SelectSingleNode("/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results", nsmgr);

            // Elementlerin oluşturulması
            XmlElement result = xml.CreateElement("dcc", "result", dcc);
            result.SetAttribute("id", "Effective Effiency" + str + "_dB");
            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement data = xml.CreateElement("dcc", "data", dcc);
            XmlElement list = xml.CreateElement("dcc", "list", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "Effective Effiency" + str + "_dB";

            //Dataları içeren elementler oluşturulup dcc:List elementine eklenir.
            list.AppendChild(add_EE_Frekans(xml, dataXml));

            if (control[0])
            {
                List<XmlElement> xmlList = Add_EE(xml, dataXml.XML_EE_ArrayEE, dataXml.XML_EE_ArrayEEUnc, "EE");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[1])
            {
                List<XmlElement> xmlList = Add_EE_Reel_Imag(xml, dataXml.XML_EE_ArrayS11Reel, dataXml.XML_EE_ArrayS11ReelUnc, dataXml.XML_EE_ArrayS11Complex, dataXml.XML_EE_ArrayS11ComplexUnc, "EE");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[2])
            {
                List<XmlElement> xmlList = Add_RHO(xml, dataXml.XML_EE_ArrayRhoLin, dataXml.XML_EE_ArrayRhoUnc, "EE");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[3])
            {
                List<XmlElement> xmlList = Add_EE_CF(xml, dataXml.XML_EE_ArrayCF, dataXml.XML_EE_ArrayCFUnc, "EE");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }


            }
            //Son eklemeler yapılarak data geçirme tamamlanır.
            name.AppendChild(content);
            result.AppendChild(name);
            data.AppendChild(list);
            result.AppendChild(data);
            sResults.AppendChild(result);

            return xml;


        }

        public XmlElement add_EE_Frekans(XmlDocument xml, XML_Arrays dataXml)
        {
            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            XmlElement frekansElement = xml.CreateElement("dcc", "quantity", dcc);
            frekansElement.SetAttribute("refType", "frequency_ee");

            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "Frequency";

            XmlElement hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement realList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement unit = xml.CreateElement("si", "unitXMLList", si);
            unit.InnerText = "\\dB";

            string frekansData = string.Join(" ", dataXml.XML_EE_ArrayFrekans.ToArray());
            value.InnerText = frekansData;

            name.AppendChild(content);
            realList.AppendChild(value);
            realList.AppendChild(unit);
            hibrid.AppendChild(realList);
            frekansElement.AppendChild(name);
            frekansElement.AppendChild(hibrid);

            return frekansElement;
        }


        public List<XmlElement> Add_EE(XmlDocument xml, ArrayList EE_Array, ArrayList EE_Unc_Array, string EE)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // EE Element Oluşturulması
            XmlElement EE_Element = xml.CreateElement("dcc", "quantity", dcc);
            EE_Element.SetAttribute("refType", "Effective Effiency " + EE + "-EE");

            XmlElement EE_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement EE_Content = xml.CreateElement("dcc", "content", dcc);
            EE_Content.InnerText = EE + " EE_";

            XmlElement EE_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement EE_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement EE_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement EE_Unit = xml.CreateElement("si", "unitXMLList", si);
            EE_Unit.InnerText = "\\dB";

            string EE_Data = string.Join(" ", EE_Array.ToArray());
            EE_Value.InnerText = EE_Data;

            EE_Name.AppendChild(EE_Content);
            EE_RealList.AppendChild(EE_Value);
            EE_RealList.AppendChild(EE_Unit);
            EE_Hibrid.AppendChild(EE_RealList);
            EE_Element.AppendChild(EE_Name);
            EE_Element.AppendChild(EE_Hibrid);

            xmlElements.Add(EE_Element);

            // EE_ Unc Element Oluşturulması
            XmlElement EE_UncElement = xml.CreateElement("dcc", "quantity", dcc);
            EE_UncElement.SetAttribute("refType", "Effective Effiency " + EE + "-EE_Unc");

            XmlElement EE_UncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement EE_UncContent = xml.CreateElement("dcc", "content", dcc);
            EE_UncContent.InnerText = EE + " EE_Unc";

            XmlElement EE_UncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement EE_UncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement EE_UncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement EE_UncUnit = xml.CreateElement("si", "unitXMLList", si);
            EE_UncUnit.InnerText = "\\dB";

            string EE_UncData = string.Join(" ", EE_Unc_Array.ToArray());
            EE_UncValue.InnerText = EE_UncData;

            EE_UncName.AppendChild(EE_UncContent);
            EE_UncRealList.AppendChild(EE_UncValue);
            EE_UncRealList.AppendChild(EE_UncUnit);
            EE_UncHibrid.AppendChild(EE_UncRealList);
            EE_UncElement.AppendChild(EE_UncName);
            EE_UncElement.AppendChild(EE_UncHibrid);

            xmlElements.Add(EE_UncElement);

            return xmlElements;
        }

        public List<XmlElement> Add_RHO(XmlDocument xml, ArrayList Rho, ArrayList Rho_Lin, string EE)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // rho Element Oluşturulması
            XmlElement Rho_Element = xml.CreateElement("dcc", "quantity", dcc);
            Rho_Element.SetAttribute("refType", "Effective Effiency " + EE + "-Rho");

            XmlElement Rho_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement Rho_Content = xml.CreateElement("dcc", "content", dcc);
            Rho_Content.InnerText = EE + " Rho";

            XmlElement Rho_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement Rho_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement Rho_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement Rho_Unit = xml.CreateElement("si", "unitXMLList", si);
            Rho_Unit.InnerText = "\\dB";

            string Rho_Data = string.Join(" ", Rho.ToArray());
            Rho_Value.InnerText = Rho_Data;

            Rho_Name.AppendChild(Rho_Content);
            Rho_RealList.AppendChild(Rho_Value);
            Rho_RealList.AppendChild(Rho_Unit);
            Rho_Hibrid.AppendChild(Rho_RealList);
            Rho_Element.AppendChild(Rho_Name);
            Rho_Element.AppendChild(Rho_Hibrid);

            xmlElements.Add(Rho_Element);

            // Rho_ Unc Element Oluşturulması
            XmlElement Rho_UncElement = xml.CreateElement("dcc", "quantity", dcc);
            Rho_UncElement.SetAttribute("refType", "Effective Effiency " + EE + "-Rho_Unc");

            XmlElement Rho_UncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement Rho_UncContent = xml.CreateElement("dcc", "content", dcc);
            Rho_UncContent.InnerText = EE + " Rho_Unc";

            XmlElement Rho_UncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement Rho_UncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement Rho_UncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement Rho_UncUnit = xml.CreateElement("si", "unitXMLList", si);
            Rho_UncUnit.InnerText = "\\dB";

            string Rho_UncData = string.Join(" ", Rho_Lin.ToArray());
            Rho_UncValue.InnerText = Rho_UncData;

            Rho_UncName.AppendChild(Rho_UncContent);
            Rho_UncRealList.AppendChild(Rho_UncValue);
            Rho_UncRealList.AppendChild(Rho_UncUnit);
            Rho_UncHibrid.AppendChild(Rho_UncRealList);
            Rho_UncElement.AppendChild(Rho_UncName);
            Rho_UncElement.AppendChild(Rho_UncHibrid);

            xmlElements.Add(Rho_UncElement);

            return xmlElements;
        }

        public List<XmlElement> Add_EE_CF(XmlDocument xml, ArrayList EE_CF_Array, ArrayList EE__CF_Unc_Array, string EE)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // EE_CF_ Element Oluşturulması
            XmlElement EE_CF_Element = xml.CreateElement("dcc", "quantity", dcc);
            EE_CF_Element.SetAttribute("refType", "Effective Effiency " + EE + "-Cal_Factor");

            XmlElement EE_CF_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement EE_CF_Content = xml.CreateElement("dcc", "content", dcc);
            EE_CF_Content.InnerText = EE + " EE_Cal_Factor";

            XmlElement EE_CF_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement EE_CF_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement EE_CF_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement EE_CF_Unit = xml.CreateElement("si", "unitXMLList", si);
            EE_CF_Unit.InnerText = "\\dB";

            string EE_CF_Data = string.Join(" ", EE_CF_Array.ToArray());
            EE_CF_Value.InnerText = EE_CF_Data;

            EE_CF_Name.AppendChild(EE_CF_Content);
            EE_CF_RealList.AppendChild(EE_CF_Value);
            EE_CF_RealList.AppendChild(EE_CF_Unit);
            EE_CF_Hibrid.AppendChild(EE_CF_RealList);
            EE_CF_Element.AppendChild(EE_CF_Name);
            EE_CF_Element.AppendChild(EE_CF_Hibrid);

            xmlElements.Add(EE_CF_Element);

            // EE_CF_ Unc Element Oluşturulması
            XmlElement EE_CF_UncElement = xml.CreateElement("dcc", "quantity", dcc);
            EE_CF_UncElement.SetAttribute("refType", "Effective Effiency " + EE + "-Cal_Factor_Unc");

            XmlElement EE_CF_UncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement EE_CF_UncContent = xml.CreateElement("dcc", "content", dcc);
            EE_CF_UncContent.InnerText = EE + " EE_Cal_Factor_Unc";

            XmlElement EE_CF_UncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement EE_CF_UncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement EE_CF_UncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement EE_CF_UncUnit = xml.CreateElement("si", "unitXMLList", si);
            EE_CF_UncUnit.InnerText = "\\dB";

            string EE_CF_UncData = string.Join(" ", EE__CF_Unc_Array.ToArray());
            EE_CF_UncValue.InnerText = EE_CF_UncData;

            EE_CF_UncName.AppendChild(EE_CF_UncContent);
            EE_CF_UncRealList.AppendChild(EE_CF_UncValue);
            EE_CF_UncRealList.AppendChild(EE_CF_UncUnit);
            EE_CF_UncHibrid.AppendChild(EE_CF_UncRealList);
            EE_CF_UncElement.AppendChild(EE_CF_UncName);
            EE_CF_UncElement.AppendChild(EE_CF_UncHibrid);

            xmlElements.Add(EE_CF_UncElement);

            return xmlElements;
        }

        public List<XmlElement> Add_EE_Reel_Imag(XmlDocument xml, ArrayList Reel, ArrayList ReelUnc, ArrayList Imag, ArrayList ImagUnc, string EE)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // Lin Element Oluşturulması
            XmlElement linElement = xml.CreateElement("dcc", "quantity", dcc);
            linElement.SetAttribute("refType", "Effective Effiency " + EE + "-Reel");

            XmlElement linName = xml.CreateElement("dcc", "name", dcc);
            XmlElement linContent = xml.CreateElement("dcc", "content", dcc);
            linContent.InnerText = EE + " Reel";

            XmlElement linHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement linRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement linValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement linUnit = xml.CreateElement("si", "unitXMLList", si);
            linUnit.InnerText = "\\dB";

            string linData = string.Join(" ", Reel.ToArray());
            linValue.InnerText = linData;

            linName.AppendChild(linContent);
            linRealList.AppendChild(linValue);
            linRealList.AppendChild(linUnit);
            linHibrid.AppendChild(linRealList);
            linElement.AppendChild(linName);
            linElement.AppendChild(linHibrid);

            xmlElements.Add(linElement);

            // Lin Unc Element Oluşturulması
            XmlElement linUncElement = xml.CreateElement("dcc", "quantity", dcc);
            linUncElement.SetAttribute("refType", "Effective Effiency " + EE + "-Reel_Unc");

            XmlElement linUncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement linUncContent = xml.CreateElement("dcc", "content", dcc);
            linUncContent.InnerText = EE + " Reel_Unc";

            XmlElement linUncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement linUncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement linUncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement linUncUnit = xml.CreateElement("si", "unitXMLList", si);
            linUncUnit.InnerText = "\\dB";

            string linUncData = string.Join(" ", ReelUnc.ToArray());
            linUncValue.InnerText = linUncData;

            linUncName.AppendChild(linUncContent);
            linUncRealList.AppendChild(linUncValue);
            linUncRealList.AppendChild(linUncUnit);
            linUncHibrid.AppendChild(linUncRealList);
            linUncElement.AppendChild(linUncName);
            linUncElement.AppendChild(linUncHibrid);

            xmlElements.Add(linUncElement);

            // Phase Element Oluşturulması
            XmlElement phaseElement = xml.CreateElement("dcc", "quantity", dcc);
            phaseElement.SetAttribute("refType", "Effective Effiency " + EE + "-Imaginer");

            XmlElement phaseName = xml.CreateElement("dcc", "name", dcc);
            XmlElement phaseContent = xml.CreateElement("dcc", "content", dcc);
            phaseContent.InnerText = EE + " Imaginer";

            XmlElement phaseHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement phaseRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement phaseValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement phaseUnit = xml.CreateElement("si", "unitXMLList", si);
            phaseUnit.InnerText = "\\dB";

            string phaseData = string.Join(" ", Imag.ToArray());
            phaseValue.InnerText = phaseData;

            phaseName.AppendChild(phaseContent);
            phaseRealList.AppendChild(phaseValue);
            phaseRealList.AppendChild(phaseUnit);
            phaseHibrid.AppendChild(phaseRealList);
            phaseElement.AppendChild(phaseName);
            phaseElement.AppendChild(phaseHibrid);

            xmlElements.Add(phaseElement);

            // phase Unc Element Oluşturulması
            XmlElement phaseUncElement = xml.CreateElement("dcc", "quantity", dcc);
            phaseUncElement.SetAttribute("refType", "Effective Effiency " + EE + "-Imaginer_Unc");

            XmlElement phaseUncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement phaseUncContent = xml.CreateElement("dcc", "content", dcc);
            phaseUncContent.InnerText = EE + " Imaginer_Unc";

            XmlElement phaseUncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement phaseUncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement phaseUncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement phaseUncUnit = xml.CreateElement("si", "unitXMLList", si);
            phaseUncUnit.InnerText = "\\dB";

            string phaseUncData = string.Join(" ", ImagUnc.ToArray());
            phaseUncValue.InnerText = phaseUncData;

            phaseUncName.AppendChild(phaseUncContent);
            phaseUncRealList.AppendChild(phaseUncValue);
            phaseUncRealList.AppendChild(phaseUncUnit);
            phaseUncHibrid.AppendChild(phaseUncRealList);
            phaseUncElement.AppendChild(phaseUncName);
            phaseUncElement.AppendChild(phaseUncHibrid);

            xmlElements.Add(phaseUncElement);

            return xmlElements;
        }



        #endregion

        #region CF XML İŞLEMLERİ
        #region CF_Result
        // S Parametre için results elementi altına result elemeti oluşturma.
        public XmlDocument AddCFResult(XmlDocument xml, string str, XML_Arrays dataXml, List<bool> control)
        {
            //Result Namespace oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", dcc);
            nsmgr.AddNamespace("si", si);

            this.dataList = control;

            XmlNode sResults = xml.SelectSingleNode("/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results", nsmgr);

            // Elementlerin oluşturulması
            XmlElement result = xml.CreateElement("dcc", "result", dcc);
            result.SetAttribute("id", "Calibration_Factor" + str + "_dB");
            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement data = xml.CreateElement("dcc", "data", dcc);
            XmlElement list = xml.CreateElement("dcc", "list", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "Calibration Factor" + str + "_dB";

            //Dataları içeren elementler oluşturulup dcc:List elementine eklenir.
            list.AppendChild(addCFFrekans(xml, dataXml));

            if (control[0])
            {
                List<XmlElement> xmlList = AddCF(xml, dataXml.XML_CF_Array, dataXml.XML_CF_ArrayCFUnc, "CF");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[1])
            {
                List<XmlElement> xmlList = AddCFReelImagRef(xml, dataXml.XML_CF_ArrayReel, dataXml.XML_CF_ArrayReelUnc, dataXml.XML_CF_ArrayComplex, dataXml.XML_CF_ArrayComplexUnc, dataXml.XML_CF_YK, dataXml.XML_CF_YK_Unc, "CF");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }



            //Son eklemeler yapılarak data geçirme tamamlanır.
            name.AppendChild(content);
            result.AppendChild(name);
            data.AppendChild(list);
            result.AppendChild(data);
            sResults.AppendChild(result);

            return xml;
        }
        #endregion

        #region CF_Frequency
        public XmlElement addCFFrekans(XmlDocument xml, XML_Arrays dataXml)
        {
            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            XmlElement frekansElement = xml.CreateElement("dcc", "quantity", dcc);
            frekansElement.SetAttribute("refType", "frequency_cf");

            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "Frequency";

            XmlElement hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement realList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement unit = xml.CreateElement("si", "unitXMLList", si);
            unit.InnerText = "\\dB";

            string frekansData = string.Join(" ", dataXml.XML_CF_ArrayFrekans.ToArray());
            value.InnerText = frekansData;

            name.AppendChild(content);
            realList.AppendChild(value);
            realList.AppendChild(unit);
            hibrid.AppendChild(realList);
            frekansElement.AppendChild(name);
            frekansElement.AppendChild(hibrid);

            return frekansElement;

        }
        #endregion

        #region CF_Parameter
        public List<XmlElement> AddCF(XmlDocument xml, ArrayList cf, ArrayList cfunc, string cfstr)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // Cfrel Element Oluşturulması
            XmlElement cfrelElement = xml.CreateElement("dcc", "quantity", dcc);
            cfrelElement.SetAttribute("refType", "Calibration Factor " + cfstr + "-Cal_Factor");

            XmlElement cfrelName = xml.CreateElement("dcc", "name", dcc);
            XmlElement cfrelContent = xml.CreateElement("dcc", "content", dcc);
            cfrelContent.InnerText = cfstr;

            XmlElement cfrelHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement cfrelRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement cfrelValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement cfrelUnit = xml.CreateElement("si", "unitXMLList", si);
            cfrelUnit.InnerText = "\\dB";

            string reelData = string.Join(" ", cf.ToArray());
            cfrelValue.InnerText = reelData;

            cfrelName.AppendChild(cfrelContent);
            cfrelRealList.AppendChild(cfrelValue);
            cfrelRealList.AppendChild(cfrelUnit);
            cfrelHibrid.AppendChild(cfrelRealList);
            cfrelElement.AppendChild(cfrelName);
            cfrelElement.AppendChild(cfrelHibrid);

            xmlElements.Add(cfrelElement);

            // Cfrel Unc Element Oluşturulması
            XmlElement cfreluncElement = xml.CreateElement("dcc", "quantity", dcc);
            cfreluncElement.SetAttribute("refType", "Calibration Factor " + cfstr + "-Cal_Factor_Unc");

            XmlElement cfreluncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement cfreluncContent = xml.CreateElement("dcc", "content", dcc);
            cfreluncContent.InnerText = cfstr + " Unc";

            XmlElement cfreluncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement cfreluncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement cfreluncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement cfreluncUnit = xml.CreateElement("si", "unitXMLList", si);
            cfreluncUnit.InnerText = "\\dB";

            string cfreluncData = string.Join(" ", cfunc.ToArray());
            cfreluncValue.InnerText = cfreluncData;

            cfreluncName.AppendChild(cfreluncContent);
            cfreluncRealList.AppendChild(cfreluncValue);
            cfreluncRealList.AppendChild(cfreluncUnit);
            cfreluncHibrid.AppendChild(cfreluncRealList);
            cfreluncElement.AppendChild(cfreluncName);
            cfreluncElement.AppendChild(cfreluncHibrid);

            xmlElements.Add(cfreluncElement);
            return xmlElements;

        }

        public List<XmlElement> AddCFReelImagRef(XmlDocument xml, ArrayList cfreel, ArrayList cfreelunc, ArrayList cfimag, ArrayList cfimagunc, ArrayList cfref, ArrayList cfrefunc, string cfstr)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // Cfreel Element Oluşturulması
            XmlElement cfreelElement = xml.CreateElement("dcc", "quantity", dcc);
            cfreelElement.SetAttribute("refType", "Calibration Factor " + cfstr + "-Reel");

            XmlElement cfreelName = xml.CreateElement("dcc", "name", dcc);
            XmlElement cfreelContent = xml.CreateElement("dcc", "content", dcc);
            cfreelContent.InnerText = cfstr + " Reel";

            XmlElement cfreelHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement cfreelRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement cfreelValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement cfreelUnit = xml.CreateElement("si", "unitXMLList", si);
            cfreelUnit.InnerText = "\\dB";

            string cfreelData = string.Join(" ", cfreel.ToArray());
            cfreelValue.InnerText = cfreelData;

            cfreelName.AppendChild(cfreelContent);
            cfreelRealList.AppendChild(cfreelValue);
            cfreelRealList.AppendChild(cfreelUnit);
            cfreelHibrid.AppendChild(cfreelRealList);
            cfreelElement.AppendChild(cfreelName);
            cfreelElement.AppendChild(cfreelHibrid);

            xmlElements.Add(cfreelElement);

            // Cfreel Unc Element Oluşturulması
            XmlElement cfreeluncElement = xml.CreateElement("dcc", "quantity", dcc);
            cfreeluncElement.SetAttribute("refType", "Calibration Factor " + cfstr + "-Reel_Unc");

            XmlElement cfreeluncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement cfreeluncContent = xml.CreateElement("dcc", "content", dcc);
            cfreeluncContent.InnerText = cfstr + " ReelUnc";

            XmlElement cfreeluncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement cfreeluncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement cfreeluncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement cfreeluncUnit = xml.CreateElement("si", "unitXMLList", si);
            cfreeluncUnit.InnerText = "\\dB";

            string cfreeluncData = string.Join(" ", cfreelunc.ToArray());
            cfreeluncValue.InnerText = cfreeluncData;

            cfreeluncName.AppendChild(cfreeluncContent);
            cfreeluncRealList.AppendChild(cfreeluncValue);
            cfreeluncRealList.AppendChild(cfreeluncUnit);
            cfreeluncHibrid.AppendChild(cfreeluncRealList);
            cfreeluncElement.AppendChild(cfreeluncName);
            cfreeluncElement.AppendChild(cfreeluncHibrid);

            xmlElements.Add(cfreeluncElement);

            // Cfimag Element Oluşturulması
            XmlElement cfimagElement = xml.CreateElement("dcc", "quantity", dcc);
            cfimagElement.SetAttribute("refType", "Calibration Factor " + cfstr + "-Imaginer");

            XmlElement cfimagName = xml.CreateElement("dcc", "name", dcc);
            XmlElement cfimagContent = xml.CreateElement("dcc", "content", dcc);
            cfimagContent.InnerText = cfstr + " Imaginer";

            XmlElement cfimagHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement cfimagRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement cfimagValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement cfimagUnit = xml.CreateElement("si", "unitXMLList", si);
            cfimagUnit.InnerText = "\\dB";

            string cfimagData = string.Join(" ", cfimag.ToArray());
            cfimagValue.InnerText = cfimagData;

            cfimagName.AppendChild(cfimagContent);
            cfimagRealList.AppendChild(cfimagValue);
            cfimagRealList.AppendChild(cfimagUnit);
            cfimagHibrid.AppendChild(cfimagRealList);
            cfimagElement.AppendChild(cfimagName);
            cfimagElement.AppendChild(cfimagHibrid);

            xmlElements.Add(cfimagElement);

            // Cfimag Unc Element Oluşturulması
            XmlElement cfimaguncElement = xml.CreateElement("dcc", "quantity", dcc);
            cfimaguncElement.SetAttribute("refType", "Calibration Factor " + cfstr + "-Imaginer_Unc");

            XmlElement cfimaguncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement cfimaguncContent = xml.CreateElement("dcc", "content", dcc);
            cfimaguncContent.InnerText = cfstr + " ImaginerUnc";

            XmlElement cfimaguncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement cfimaguncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement cfimaguncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement cfimaguncUnit = xml.CreateElement("si", "unitXMLList", si);
            cfimaguncUnit.InnerText = "\\dB";

            string cfimaguncData = string.Join(" ", cfimagunc.ToArray());
            cfimaguncValue.InnerText = cfimaguncData;

            cfimaguncName.AppendChild(cfimaguncContent);
            cfimaguncRealList.AppendChild(cfimaguncValue);
            cfimaguncRealList.AppendChild(cfimaguncUnit);
            cfimaguncHibrid.AppendChild(cfimaguncRealList);
            cfimaguncElement.AppendChild(cfimaguncName);
            cfimaguncElement.AppendChild(cfimaguncHibrid);

            xmlElements.Add(cfimaguncElement);



            // Cfref Element Oluşturulması
            XmlElement cfrefElement = xml.CreateElement("dcc", "quantity", dcc);
            cfrefElement.SetAttribute("refType", "Calibration Factor " + cfstr + "-ReflectionCof");

            XmlElement cfrefName = xml.CreateElement("dcc", "name", dcc);
            XmlElement cfrefContent = xml.CreateElement("dcc", "content", dcc);
            cfrefContent.InnerText = cfstr + " ReflectionCof";

            XmlElement cfrefHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement cfrefRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement cfrefValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement cfrefUnit = xml.CreateElement("si", "unitXMLList", si);
            cfrefUnit.InnerText = "\\dB";

            string cfrefData = string.Join(" ", cfref.ToArray());
            cfrefValue.InnerText = cfrefData;

            cfrefName.AppendChild(cfrefContent);
            cfrefRealList.AppendChild(cfrefValue);
            cfrefRealList.AppendChild(cfrefUnit);
            cfrefHibrid.AppendChild(cfrefRealList);
            cfrefElement.AppendChild(cfrefName);
            cfrefElement.AppendChild(cfrefHibrid);

            xmlElements.Add(cfrefElement);



            // Cfrefunc Element Oluşturulması
            XmlElement cfrefuncElement = xml.CreateElement("dcc", "quantity", dcc);
            cfrefuncElement.SetAttribute("refType", "Calibration Factor " + cfstr + "-ReflectionCof_Unc");

            XmlElement cfrefuncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement cfrefuncContent = xml.CreateElement("dcc", "content", dcc);
            cfrefuncContent.InnerText = cfstr + " ReflectionCof_Unc";

            XmlElement cfrefuncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement cfrefuncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement cfrefuncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement cfrefuncUnit = xml.CreateElement("si", "unitXMLList", si);
            cfrefuncUnit.InnerText = "\\dB";

            string cfrefuncData = string.Join(" ", cfrefunc.ToArray());
            cfrefuncValue.InnerText = cfrefuncData;

            cfrefuncName.AppendChild(cfrefuncContent);
            cfrefuncRealList.AppendChild(cfrefuncValue);
            cfrefuncRealList.AppendChild(cfrefuncUnit);
            cfrefuncHibrid.AppendChild(cfrefuncRealList);
            cfrefuncElement.AppendChild(cfrefuncName);
            cfrefuncElement.AppendChild(cfrefuncHibrid);

            xmlElements.Add(cfrefuncElement);

            return xmlElements;

        }


        #endregion

        #endregion

        #region CIS XML İŞLEMLERİ
        public XmlDocument AddCISResult(XmlDocument xml, string str, XML_Arrays dataXml, bool control)
        {
            //Result Namespace oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", dcc);
            nsmgr.AddNamespace("si", si);

            this.CIS_bool = control;

            XmlNode sResults = xml.SelectSingleNode("/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results", nsmgr);

            // Elementlerin oluşturulması
            XmlElement result = xml.CreateElement("dcc", "result", dcc);
            result.SetAttribute("id", "Calculable Impedance Standard" + str + "_dB");
            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement data = xml.CreateElement("dcc", "data", dcc);
            XmlElement list = xml.CreateElement("dcc", "list", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "Calculable Impedance Standard" + str + "_dB";

            //Dataları içeren elementler oluşturulup dcc:List elementine eklenir.
            list.AppendChild(addCISFrekans(xml, dataXml));

            if (control)
            {
                List<XmlElement> xmlList = AddCIS_ZPosition(xml, dataXml.XML_CIS_ZP, dataXml.XML_CIS_ZP_Unc, dataXml.XML_CIS_ICOD, dataXml.XML_CIS_ICOD_Unc, dataXml.XML_CIS_OCID, dataXml.XML_CIS_OCID_Unc, "CIS");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }

            //Son eklemeler yapılarak data geçirme tamamlanır.
            name.AppendChild(content);
            result.AppendChild(name);
            data.AppendChild(list);
            result.AppendChild(data);
            sResults.AppendChild(result);

            return xml;


        }


        public XmlElement addCISFrekans(XmlDocument xml, XML_Arrays dataXml)
        {
            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            XmlElement frekansElement = xml.CreateElement("dcc", "quantity", dcc);
            frekansElement.SetAttribute("refType", "frequency_cis");

            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "Frequency";

            XmlElement hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement realList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement unit = xml.CreateElement("si", "unitXMLList", si);
            unit.InnerText = "\\dB";

            string frekansData = string.Join(" ", dataXml.XML_CIS_Olcum_Adım.ToArray());
            value.InnerText = frekansData;

            name.AppendChild(content);
            realList.AppendChild(value);
            realList.AppendChild(unit);
            hibrid.AppendChild(realList);
            frekansElement.AppendChild(name);
            frekansElement.AppendChild(hibrid);

            return frekansElement;

        }

        public List<XmlElement> AddCIS_ZPosition(XmlDocument xml, ArrayList Z_Position, ArrayList Z_Position_Unc, ArrayList ICOD, ArrayList ICOD_Unc, ArrayList OCID, ArrayList OCID_Unc, string CIS)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // CIS Z-POSİTİON  Element Oluşturulması
            XmlElement CIS_ZP_Element = xml.CreateElement("dcc", "quantity", dcc);
            CIS_ZP_Element.SetAttribute("refType", "Calculable Impedance Standard " + CIS + "-Z-Position");

            XmlElement CIS_ZP_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement CIS_ZP_Content = xml.CreateElement("dcc", "content", dcc);
            CIS_ZP_Content.InnerText = CIS + " Z-Position ";

            XmlElement CIS_ZP_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement CIS_ZP_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement CIS_ZP_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement CIS_ZP_Unit = xml.CreateElement("si", "unitXMLList", si);
            CIS_ZP_Unit.InnerText = "\\dB";

            string CIS_ZP_Data = string.Join(" ", Z_Position.ToArray());
            CIS_ZP_Value.InnerText = CIS_ZP_Data;

            CIS_ZP_Name.AppendChild(CIS_ZP_Content);
            CIS_ZP_RealList.AppendChild(CIS_ZP_Value);
            CIS_ZP_RealList.AppendChild(CIS_ZP_Unit);
            CIS_ZP_Hibrid.AppendChild(CIS_ZP_RealList);
            CIS_ZP_Element.AppendChild(CIS_ZP_Name);
            CIS_ZP_Element.AppendChild(CIS_ZP_Hibrid);

            xmlElements.Add(CIS_ZP_Element);

            //  CIS Z-POSİTİON UNC  Element Oluşturulması
            XmlElement CIS_ZP_Unc_Element = xml.CreateElement("dcc", "quantity", dcc);
            CIS_ZP_Unc_Element.SetAttribute("refType", "Calculable Impedance Standard " + CIS + "-Z-PositionUnc");

            XmlElement CIS_ZP_Unc_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement CIS_ZP_Unc_Content = xml.CreateElement("dcc", "content", dcc);
            CIS_ZP_Unc_Content.InnerText = CIS + " Z-Position Unc";

            XmlElement CIS_ZP_Unc_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement CIS_ZP_Unc_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement CIS_ZP_Unc_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement CIS_ZP_Unc_Unit = xml.CreateElement("si", "unitXMLList", si);
            CIS_ZP_Unc_Unit.InnerText = "\\dB";

            string CIS_ZP_Unc_Data = string.Join(" ", Z_Position_Unc.ToArray());
            CIS_ZP_Unc_Value.InnerText = CIS_ZP_Unc_Data;

            CIS_ZP_Unc_Name.AppendChild(CIS_ZP_Unc_Content);
            CIS_ZP_Unc_RealList.AppendChild(CIS_ZP_Unc_Value);
            CIS_ZP_Unc_RealList.AppendChild(CIS_ZP_Unc_Unit);
            CIS_ZP_Unc_Hibrid.AppendChild(CIS_ZP_Unc_RealList);
            CIS_ZP_Unc_Element.AppendChild(CIS_ZP_Unc_Name);
            CIS_ZP_Unc_Element.AppendChild(CIS_ZP_Unc_Hibrid);

            xmlElements.Add(CIS_ZP_Unc_Element);

            // CIS ICOD 
            XmlElement ICOD_Element = xml.CreateElement("dcc", "quantity", dcc);
            ICOD_Element.SetAttribute("refType", "Calculable Impedance Standard " + CIS + "-ICOD");

            XmlElement ICOD_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement ICOD_content = xml.CreateElement("dcc", "content", dcc);
            ICOD_content.InnerText = CIS + " ICOD ";

            XmlElement ICOD_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement ICOD_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement ICOD_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement ICOD_Unit = xml.CreateElement("si", "unitXMLList", si);
            ICOD_Unit.InnerText = "\\dB";

            string ICOD_Data = string.Join(" ", ICOD.ToArray());
            ICOD_Value.InnerText = ICOD_Data;

            ICOD_Name.AppendChild(ICOD_content);
            ICOD_RealList.AppendChild(ICOD_Value);
            ICOD_RealList.AppendChild(ICOD_Unit);
            ICOD_Hibrid.AppendChild(ICOD_RealList);
            ICOD_Element.AppendChild(ICOD_Name);
            ICOD_Element.AppendChild(ICOD_Hibrid);

            xmlElements.Add(ICOD_Element);

            // ICOD  UNC 
            XmlElement ICOD_Unc_Element = xml.CreateElement("dcc", "quantity", dcc);
            ICOD_Unc_Element.SetAttribute("refType", "Calculable Impedance Standard " + CIS + "-ICOD_Unc");

            XmlElement ICOD_Unc_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement ICOD_Unc_Content = xml.CreateElement("dcc", "content", dcc);
            ICOD_Unc_Content.InnerText = CIS + " ICOD Unc";

            XmlElement ICOD_Unc_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement ICOD_Unc_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement ICOD_Unc_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement ICOD_Unc_Unit = xml.CreateElement("si", "unitXMLList", si);
            ICOD_Unc_Unit.InnerText = "\\dB";

            string ICOD_Unc_Data = string.Join(" ", ICOD_Unc.ToArray());
            ICOD_Unc_Value.InnerText = ICOD_Unc_Data;

            ICOD_Unc_Name.AppendChild(ICOD_Unc_Content);
            ICOD_Unc_RealList.AppendChild(ICOD_Unc_Value);
            ICOD_Unc_RealList.AppendChild(ICOD_Unc_Unit);
            ICOD_Unc_Hibrid.AppendChild(ICOD_Unc_RealList);
            ICOD_Unc_Element.AppendChild(ICOD_Unc_Name);
            ICOD_Unc_Element.AppendChild(ICOD_Unc_Hibrid);

            xmlElements.Add(ICOD_Unc_Element);



            // OCID  
            XmlElement OCID_Element = xml.CreateElement("dcc", "quantity", dcc);
            OCID_Element.SetAttribute("refType", "Calculable Impedance Standard " + CIS + "-OCID");

            XmlElement OCID_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement OCID_Content = xml.CreateElement("dcc", "content", dcc);
            OCID_Content.InnerText = CIS + " OCID  ";

            XmlElement OCID_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement OCID_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement OCID_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement OCID_Unit = xml.CreateElement("si", "unitXMLList", si);
            OCID_Unit.InnerText = "\\dB";

            string OCID_Data = string.Join(" ", OCID.ToArray());
            OCID_Value.InnerText = OCID_Data;

            OCID_Name.AppendChild(OCID_Content);
            OCID_RealList.AppendChild(OCID_Value);
            OCID_RealList.AppendChild(OCID_Unit);
            OCID_Hibrid.AppendChild(OCID_RealList);
            OCID_Element.AppendChild(OCID_Name);
            OCID_Element.AppendChild(OCID_Hibrid);

            xmlElements.Add(OCID_Element);


            // OCID  Unc Element Oluşturulması
            XmlElement OCID_Unc_Element = xml.CreateElement("dcc", "quantity", dcc);
            OCID_Unc_Element.SetAttribute("refType", "Calculable Impedance Standard  " + CIS + "-OCID_Unc");

            XmlElement OCID_Unc_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement OCID_Unc_Content = xml.CreateElement("dcc", "content", dcc);
            OCID_Unc_Content.InnerText = CIS + " OCID Unc";

            XmlElement OCID_Unc_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement OCID_Unc_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement OCID_Unc_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement OCID_Unc_Unit = xml.CreateElement("si", "unitXMLList", si);
            OCID_Unc_Unit.InnerText = "\\dB";

            string OCID_Unc_Data = string.Join(" ", OCID_Unc.ToArray());
            OCID_Unc_Value.InnerText = OCID_Unc_Data;

            OCID_Unc_Name.AppendChild(OCID_Unc_Content);
            OCID_Unc_RealList.AppendChild(OCID_Unc_Value);
            OCID_Unc_RealList.AppendChild(OCID_Unc_Unit);
            OCID_Unc_Hibrid.AppendChild(OCID_Unc_RealList);
            OCID_Unc_Element.AppendChild(OCID_Unc_Name);
            OCID_Unc_Element.AppendChild(OCID_Unc_Hibrid);

            xmlElements.Add(OCID_Unc_Element);

            return xmlElements;

        }
        #endregion

        #region NOİSE XML İŞLEMLERİ

        #region Noise_Result

        public XmlDocument AddNoiseResult(XmlDocument xml, string str, XML_Arrays dataXml, List<bool> control)
        {
            //Result Namespace oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", dcc);
            nsmgr.AddNamespace("si", si);

            this.dataList = control;

            XmlNode sResults = xml.SelectSingleNode("/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results", nsmgr);

            // Elementlerin oluşturulması
            XmlElement result = xml.CreateElement("dcc", "result", dcc);
            result.SetAttribute("id", "Noise" + str + "_dB");
            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement data = xml.CreateElement("dcc", "data", dcc);
            XmlElement list = xml.CreateElement("dcc", "list", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "Noise" + str + "_dB";

            //Dataları içeren elementler oluşturulup dcc:List elementine eklenir.
            list.AppendChild(addNoiseFrekans(xml, dataXml));

            if (control[0])
            {
                List<XmlElement> xmlList = AddENR(xml, dataXml.XML_NS_ArrayENR, dataXml.XML_NS_ArrayENRUnc, "Noise");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }

            if (control[1])
            {
                List<XmlElement> xmlList = Add_RC_CON_OFF(xml, dataXml.XML_NS_ArrayRC, dataXml.XML_NS_ArrayRC_ustlimit, dataXml.XML_NS_ArrayRCUnc, dataXml.XML_NS_ArrayRC_Phase, dataXml.XML_NS_ArrayRC_PhaseUnc, "Noise_DC_ON");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[2])
            {
                List<XmlElement> xmlList = Add_RC_CON_OFF(xml, dataXml.XML_NS_ArrayRC_DC_OFF, dataXml.XML_NS_ArrayRC_ustlimit_DC_OFF, dataXml.XML_NS_ArrayRCUnc_DC_OFF, dataXml.XML_NS_ArrayRC_Phase_DC_OFF, dataXml.XML_NS_ArrayRC_PhaseUnc_DC_OFF, "Noise_DC_OFF");

                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }



            //Son eklemeler yapılarak data geçirme tamamlanır.
            name.AppendChild(content);
            result.AppendChild(name);
            data.AppendChild(list);
            result.AppendChild(data);
            sResults.AppendChild(result);

            return xml;
        }
        #endregion

        #region Noise_Frequency
        public XmlElement addNoiseFrekans(XmlDocument xml, XML_Arrays dataXml)
        {
            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            XmlElement frekansElement = xml.CreateElement("dcc", "quantity", dcc);
            frekansElement.SetAttribute("refType", "frequency_noise");

            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "Frequency";

            XmlElement hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement realList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement unit = xml.CreateElement("si", "unitXMLList", si);
            unit.InnerText = "\\dB";

            string frekansData = string.Join(" ", dataXml.XML_NS_ArrayFrekans.ToArray());
            value.InnerText = frekansData;

            name.AppendChild(content);
            realList.AppendChild(value);
            realList.AppendChild(unit);
            hibrid.AppendChild(realList);
            frekansElement.AppendChild(name);
            frekansElement.AppendChild(hibrid);

            return frekansElement;

        }
        #endregion

        #region Noise_Parameter
        public List<XmlElement> AddENR(XmlDocument xml, ArrayList enr, ArrayList uenr, string noisestr)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // enr Element Oluşturulması
            XmlElement enrElement = xml.CreateElement("dcc", "quantity", dcc);
            enrElement.SetAttribute("refType", noisestr + "_ENR");

            XmlElement enrName = xml.CreateElement("dcc", "name", dcc);
            XmlElement enrContent = xml.CreateElement("dcc", "content", dcc);
            enrContent.InnerText = noisestr + "_ENR";

            XmlElement enrHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement enrRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement enrValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement enrUnit = xml.CreateElement("si", "unitXMLList", si);
            enrUnit.InnerText = "\\dB";

            string reelData = string.Join(" ", enr.ToArray());
            enrValue.InnerText = reelData;

            enrName.AppendChild(enrContent);
            enrRealList.AppendChild(enrValue);
            enrRealList.AppendChild(enrUnit);
            enrHibrid.AppendChild(enrRealList);
            enrElement.AppendChild(enrName);
            enrElement.AppendChild(enrHibrid);

            xmlElements.Add(enrElement);

            // Uenr Element Oluşturulması
            XmlElement uenrElement = xml.CreateElement("dcc", "quantity", dcc);
            uenrElement.SetAttribute("refType", noisestr + "_ENR_Unc");

            XmlElement uenrName = xml.CreateElement("dcc", "name", dcc);
            XmlElement uenrContent = xml.CreateElement("dcc", "content", dcc);
            uenrContent.InnerText = noisestr + "_ENR_Unc ";

            XmlElement uenrHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement uenrRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement uenrValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement uenrUnit = xml.CreateElement("si", "unitXMLList", si);
            uenrUnit.InnerText = "\\dB";

            string uenrData = string.Join(" ", uenr.ToArray());
            uenrValue.InnerText = uenrData;

            uenrName.AppendChild(uenrContent);
            uenrRealList.AppendChild(uenrValue);
            uenrRealList.AppendChild(uenrUnit);
            uenrHibrid.AppendChild(uenrRealList);
            uenrElement.AppendChild(uenrName);
            uenrElement.AppendChild(uenrHibrid);

            xmlElements.Add(uenrElement);
            return xmlElements;

        }

        public List<XmlElement> Add_RC_CON_OFF(XmlDocument xml, ArrayList rclin, ArrayList rclimit, ArrayList rclinunc, ArrayList rcphase, ArrayList rcphaseunc, string noisestr)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // lin Element Oluşturulması
            XmlElement rclinElement = xml.CreateElement("dcc", "quantity", dcc);
            rclinElement.SetAttribute("refType", noisestr + "_Lin");

            XmlElement rclinName = xml.CreateElement("dcc", "name", dcc);
            XmlElement rclinContent = xml.CreateElement("dcc", "content", dcc);
            rclinContent.InnerText = noisestr + "_Lin";

            XmlElement rclinHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement rclinRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement rclinValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement rclinUnit = xml.CreateElement("si", "unitXMLList", si);
            rclinUnit.InnerText = "\\dB";

            string rclinData = string.Join(" ", rclin.ToArray());
            rclinValue.InnerText = rclinData;

            rclinName.AppendChild(rclinContent);
            rclinRealList.AppendChild(rclinValue);
            rclinRealList.AppendChild(rclinUnit);
            rclinHibrid.AppendChild(rclinRealList);
            rclinElement.AppendChild(rclinName);
            rclinElement.AppendChild(rclinHibrid);

            xmlElements.Add(rclinElement);

            // rclimit Element Oluşturulması
            XmlElement rclimitElement = xml.CreateElement("dcc", "quantity", dcc);
            rclimitElement.SetAttribute("refType", noisestr + "_Limit");

            XmlElement rclimitName = xml.CreateElement("dcc", "name", dcc);
            XmlElement rclimitContent = xml.CreateElement("dcc", "content", dcc);
            rclimitContent.InnerText = noisestr + "_Limit";

            XmlElement rclimitHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement rclimitRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement rclimitValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement rclimitUnit = xml.CreateElement("si", "unitXMLList", si);
            rclimitUnit.InnerText = "\\dB";

            string rclimitData = string.Join(" ", rclimit.ToArray());
            rclimitValue.InnerText = rclimitData;

            rclimitName.AppendChild(rclimitContent);
            rclimitRealList.AppendChild(rclimitValue);
            rclimitRealList.AppendChild(rclimitUnit);
            rclimitHibrid.AppendChild(rclimitRealList);
            rclimitElement.AppendChild(rclimitName);
            rclimitElement.AppendChild(rclimitHibrid);

            xmlElements.Add(rclimitElement);

            // lin unc Element Oluşturulması
            XmlElement rclinuncElement = xml.CreateElement("dcc", "quantity", dcc);
            rclinuncElement.SetAttribute("refType", noisestr + "_Lin_Unc");

            XmlElement rclinuncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement rclinuncContent = xml.CreateElement("dcc", "content", dcc);
            rclinuncContent.InnerText = noisestr + "_Lin_Unc";

            XmlElement rclinuncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement rclinuncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement rclinuncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement rclinuncUnit = xml.CreateElement("si", "unitXMLList", si);
            rclinuncUnit.InnerText = "\\dB";

            string rclinuncData = string.Join(" ", rclinunc.ToArray());
            rclinuncValue.InnerText = rclinuncData;

            rclinuncName.AppendChild(rclinuncContent);
            rclinuncRealList.AppendChild(rclinuncValue);
            rclinuncRealList.AppendChild(rclinuncUnit);
            rclinuncHibrid.AppendChild(rclinuncRealList);
            rclinuncElement.AppendChild(rclinuncName);
            rclinuncElement.AppendChild(rclinuncHibrid);

            xmlElements.Add(rclinuncElement);

            // rcphase Element Oluşturulması
            XmlElement rcphaseElement = xml.CreateElement("dcc", "quantity", dcc);
            rcphaseElement.SetAttribute("refType", noisestr + "_RC_Phase");

            XmlElement rcphaseName = xml.CreateElement("dcc", "name", dcc);
            XmlElement rcphaseContent = xml.CreateElement("dcc", "content", dcc);
            rcphaseContent.InnerText = noisestr + "_RC_Phase";

            XmlElement rcphaseHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement rcphaseRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement rcphaseValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement rcphaseUnit = xml.CreateElement("si", "unitXMLList", si);
            rcphaseUnit.InnerText = "\\dB";

            string rcphaseData = string.Join(" ", rcphase.ToArray());
            rcphaseValue.InnerText = rcphaseData;

            rcphaseName.AppendChild(rcphaseContent);
            rcphaseRealList.AppendChild(rcphaseValue);
            rcphaseRealList.AppendChild(rcphaseUnit);
            rcphaseHibrid.AppendChild(rcphaseRealList);
            rcphaseElement.AppendChild(rcphaseName);
            rcphaseElement.AppendChild(rcphaseHibrid);

            xmlElements.Add(rcphaseElement);

            // rcphaseunc Element Oluşturulması
            XmlElement rcphaseuncElement = xml.CreateElement("dcc", "quantity", dcc);
            rcphaseuncElement.SetAttribute("refType", noisestr + "_RC_Phase_Unc");

            XmlElement rcphaseuncName = xml.CreateElement("dcc", "name", dcc);
            XmlElement rcphaseuncContent = xml.CreateElement("dcc", "content", dcc);
            rcphaseuncContent.InnerText = noisestr + "_RC_Phase_Unc";

            XmlElement rcphaseuncHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement rcphaseuncRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement rcphaseuncValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement rcphaseuncUnit = xml.CreateElement("si", "unitXMLList", si);
            rcphaseuncUnit.InnerText = "\\dB";

            string rcphaseuncData = string.Join(" ", rcphaseunc.ToArray());
            rcphaseuncValue.InnerText = rcphaseuncData;

            rcphaseuncName.AppendChild(rcphaseuncContent);
            rcphaseuncRealList.AppendChild(rcphaseuncValue);
            rcphaseuncRealList.AppendChild(rcphaseuncUnit);
            rcphaseuncHibrid.AppendChild(rcphaseuncRealList);
            rcphaseuncElement.AppendChild(rcphaseuncName);
            rcphaseuncElement.AppendChild(rcphaseuncHibrid);

            xmlElements.Add(rcphaseuncElement);



            return xmlElements;

        }

        #endregion

        #endregion

        #region ABSOLUTE RF POWER
        public XmlDocument Add_ARFP_Result(XmlDocument xml, string str, XML_Arrays dataXml, List<bool> control)
        {
            //Result Namespace oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", dcc);
            nsmgr.AddNamespace("si", si);

            this.dataList = control;

            XmlNode sResults = xml.SelectSingleNode("/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results", nsmgr);

            // Elementlerin oluşturulması
            XmlElement result = xml.CreateElement("dcc", "result", dcc);
            result.SetAttribute("id", "Absolute RF Power" + str + "_dB");
            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement data = xml.CreateElement("dcc", "data", dcc);
            XmlElement list = xml.CreateElement("dcc", "list", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "Absolute RF Power" + str + "_dB";

            //Dataları içeren elementler oluşturulup dcc:List elementine eklenir.


            if (control[0])
            {
                list.AppendChild(add_ARFP_Frekans(xml, dataXml.XML_ARFP_T1_Frekans, "t1"));
                List<XmlElement> xmlList = Add_ARFP_1(xml, dataXml.XML_ARFP_T1_Cıkıs_Gücü, dataXml.XML_ARFP_T1_Olculen_Güc, dataXml.XML_ARFP_T1_AltSınır, dataXml.XML_ARFP_T1_Sapma, dataXml.XML_ARFP_T1_ÜstSınır, dataXml.XML_ARFP_T1_Belirsizlik, "Abs_RF_Power", "Output_Power_t1", "Measured_Power_t1", "Lower_limit_t1", "Deflection_t1", "Upper_Limit_t1", "Uncertainty_t1");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }

            if (control[1])
            {
                list.AppendChild(add_ARFP_Frekans(xml, dataXml.XML_ARFP_T2_Frekans, "t2"));
                List<XmlElement> xmlList = Add_ARFP_1(xml, dataXml.XML_ARFP_T2_Cıkıs_Gücü, dataXml.XML_ARFP_T2_OlculenDeger, dataXml.XML_ARFP_T2_AltSınır, dataXml.XML_ARFP_T2_Sapma, dataXml.XML_ARFP_T2_ÜstSınır, dataXml.XML_ARFP_T2_Belirsizlik, "Abs_RF_Power", "Output_Power_t2", "Measured_Power_t2", "Lower_limit_t2", "difference_t2", "Upper_Limit_t2", "Uncertainty_t2");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[2])
            {
                list.AppendChild(add_ARFP_Frekans(xml, dataXml.XML_ARFP_T3_Frekans, "t3"));
                List<XmlElement> xmlList = Add_ARFP_1(xml, dataXml.XML_ARFP_T3_Cıkıs_Gücü, dataXml.XML_ARFP_T3_OlculenZayıflatma, dataXml.XML_ARFP_T3_AltSınır, dataXml.XML_ARFP_T3_Sapma, dataXml.XML_ARFP_T3_ÜstSınır, dataXml.XML_ARFP_T3_Belirsizlik, "Abs_RF_Power", "Output_Power_t3", "Measured_Attenuation_t3", "lower_Limit_t3", "Attenuation_Error_t3", "Upper_Limit_t3", "Uncertainty_t3");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[3])
            {
                list.AppendChild(add_ARFP_Frekans(xml, dataXml.XML_ARFP_T4_T5_T6_frekans, "t4"));
                List<XmlElement> xmlList = Add_ARFP_2(xml, dataXml.XML_ARFP_T4_SWR_Seviye, dataXml.XML_ARFP_T4_SWR_OlculenDeger, dataXml.XML_ARFP_T4_SWR_MaksimumDeger, dataXml.XML_ARFP_T4_SWR_Belirsizlik, "Abs_RF_Power", "Level_t4", "Measured_Value_t4", "Maximum_t4", "Uncertainty_t4");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[4])
            {
                list.AppendChild(add_ARFP_Frekans(xml, dataXml.XML_ARFP_T4_T5_T6_frekans, "t5"));
                List<XmlElement> xmlList = Add_ARFP_2(xml, dataXml.XML_ARFP_T5_SWR_Seviye, dataXml.XML_ARFP_T5_SWR_OlculenDeger, dataXml.XML_ARFP_T5_SWR_MaksimumDeger, dataXml.XML_ARFP_T5_SWR_Belirsizlik, "Abs_RF_Power", "Level_t5", "Measured_Value_t5", "Maximum_t5", "Uncertainty_t5");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[5])
            {
                list.AppendChild(add_ARFP_Frekans(xml, dataXml.XML_ARFP_T4_T5_T6_frekans, "t6"));
                List<XmlElement> xmlList = Add_ARFP_2(xml, dataXml.XML_ARFP_T6_SWR_Seviye, dataXml.XML_ARFP_T6_SWR_OlculenDeger, dataXml.XML_ARFP_T6_SWR_MaksimumDeger, dataXml.XML_ARFP_T6_SWR_Belirsizlik, "Abs_RF_Power", "Level_t6", "Measured_Value_t6", "Maximum_t6", "Uncertainty_t6");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[6])
            {
                list.AppendChild(add_ARFP_Frekans(xml, dataXml.XML_ARFP_T7_Frekans, "t7"));
                List<XmlElement> xmlList = Add_ARFP_1(xml, dataXml.XML_ARFP_T7_Cıkıs_Gücü, dataXml.XML_ARFP_T7_OlculenGuc, dataXml.XML_ARFP_T7_AltSınır, dataXml.XML_ARFP_T7_Sapma, dataXml.XML_ARFP_T7_ÜstSınır, dataXml.XML_ARFP_T7_Belirsizlik, "Abs_RF_Power", "Output_Power_t7", "Measured_Power_t7", "Lower_Limit_t7", "Deflection_t7", "Upper_Limit_t7", "Uncertainty_t7");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[7])
            {
                list.AppendChild(add_ARFP_Frekans(xml, dataXml.XML_ARFP_T8_Frekans, "t8"));
                List<XmlElement> xmlList = Add_ARFP_1(xml, dataXml.XML_ARFP_T8_Cıkıs_Gücü, dataXml.XML_ARFP_T8_OlculenDeger, dataXml.XML_ARFP_T8_AltSınır, dataXml.XML_ARFP_T8_Sapma, dataXml.XML_ARFP_T8_ÜstSınır, dataXml.XML_ARFP_T8_Belirsizlik, "Abs_RF_Power", "Output_Power_t8", "Measured_Value_t8", "Lower_Limit_t8", "Difference_t8", "Upper_Limit_t8", "Uncertainty_t8");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[8])
            {
                list.AppendChild(add_ARFP_Frekans(xml, dataXml.XML_ARFP_T9_T10_T11_frekans, "t9"));
                List<XmlElement> xmlList = Add_ARFP_2(xml, dataXml.XML_ARFP_T9_SWR_Seviye, dataXml.XML_ARFP_T9_SWR_OlculenDeger, dataXml.XML_ARFP_T9_SWR_MaksimumDeger, dataXml.XML_ARFP_T9_SWR_Belirsizlik, "Abs_RF_Power", "Level_t9", "Measured_Value_t9", "Upper_Limit_t9", "Uncertainty_t9");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[9])
            {
                list.AppendChild(add_ARFP_Frekans(xml, dataXml.XML_ARFP_T9_T10_T11_frekans, "t10"));
                List<XmlElement> xmlList = Add_ARFP_2(xml, dataXml.XML_ARFP_T10_SWR_Seviye, dataXml.XML_ARFP_T10_SWR_OlculenDeger, dataXml.XML_ARFP_T10_SWR_MaksimumDeger, dataXml.XML_ARFP_T10_SWR_Belirsizlik, "Abs_RF_Power", "Level_t10", "Measured_Value_t10", "Upper_Limit_t10", "Uncertainty_t10");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[10])
            {
                list.AppendChild(add_ARFP_Frekans(xml, dataXml.XML_ARFP_T9_T10_T11_frekans, "t11"));
                List<XmlElement> xmlList = Add_ARFP_2(xml, dataXml.XML_ARFP_T11_SWR_Seviye, dataXml.XML_ARFP_T11_SWR_OlculenDeger, dataXml.XML_ARFP_T11_SWR_MaksimumDeger, dataXml.XML_ARFP_T11_SWR_Belirsizlik, "Abs_RF_Power", "Level_t11", "Measured_Value_t11", "Upper_Limit_t11", "Uncertainty_t11");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }




            //Son eklemeler yapılarak data geçirme tamamlanır.
            name.AppendChild(content);
            result.AppendChild(name);
            data.AppendChild(list);
            result.AppendChild(data);
            sResults.AppendChild(result);

            return xml;
        }
        public XmlElement add_ARFP_Frekans(XmlDocument xml, ArrayList arrayListFrekans, string tableno)
        {
            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            XmlElement frekansElement = xml.CreateElement("dcc", "quantity", dcc);
            frekansElement.SetAttribute("refType", "frequency_ARFP" + "_" + tableno);

            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "Frequency";

            XmlElement hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement realList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement unit = xml.CreateElement("si", "unitXMLList", si);
            unit.InnerText = "\\dB";

            string frekansData = string.Join(" ", arrayListFrekans.ToArray());
            value.InnerText = frekansData;

            name.AppendChild(content);
            realList.AppendChild(value);
            realList.AppendChild(unit);
            hibrid.AppendChild(realList);
            frekansElement.AppendChild(name);
            frekansElement.AppendChild(hibrid);

            return frekansElement;

        }
        public List<XmlElement> Add_ARFP_1(XmlDocument xml, ArrayList col1, ArrayList col2, ArrayList col3, ArrayList col4, ArrayList col5, ArrayList col6, string ARFP_str, string Coltext1, string Coltext2, string Coltext3, string Coltext4, string Coltext5, string Coltext6)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Çıkış gücü 
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // lin Element Oluşturulması
            XmlElement ARFP_OP_Element = xml.CreateElement("dcc", "quantity", dcc);
            ARFP_OP_Element.SetAttribute("refType", ARFP_str + "_" + Coltext1);

            XmlElement ARFP_OP_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement ARFP_OP_Content = xml.CreateElement("dcc", "content", dcc);
            ARFP_OP_Content.InnerText = ARFP_str + "_" + Coltext1;

            XmlElement ARFP_OP_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement ARFP_OP_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement ARFP_OP_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement ARFP_OP_Unit = xml.CreateElement("si", "unitXMLList", si);
            ARFP_OP_Unit.InnerText = "\\dB";

            string ARFP_OP_Data = string.Join(" ", col1.ToArray());
            ARFP_OP_Value.InnerText = ARFP_OP_Data;

            ARFP_OP_Name.AppendChild(ARFP_OP_Content);
            ARFP_OP_RealList.AppendChild(ARFP_OP_Value);
            ARFP_OP_RealList.AppendChild(ARFP_OP_Unit);
            ARFP_OP_Hibrid.AppendChild(ARFP_OP_RealList);
            ARFP_OP_Element.AppendChild(ARFP_OP_Name);
            ARFP_OP_Element.AppendChild(ARFP_OP_Hibrid);

            xmlElements.Add(ARFP_OP_Element);



            //Olculen
            XmlElement measurentElement = xml.CreateElement("dcc", "quantity", dcc);
            measurentElement.SetAttribute("refType", ARFP_str + "_" + Coltext2);

            XmlElement measurentName = xml.CreateElement("dcc", "name", dcc);
            XmlElement measurentContent = xml.CreateElement("dcc", "content", dcc);
            measurentContent.InnerText = ARFP_str + "_" + Coltext2;

            XmlElement measurentHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement measurentRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement measurentValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement measurentUnit = xml.CreateElement("si", "unitXMLList", si);
            measurentUnit.InnerText = "\\dB";

            string measurentData = string.Join(" ", col2.ToArray());
            measurentValue.InnerText = measurentData;

            measurentName.AppendChild(measurentContent);
            measurentRealList.AppendChild(measurentValue);
            measurentRealList.AppendChild(measurentUnit);
            measurentHibrid.AppendChild(measurentRealList);
            measurentElement.AppendChild(measurentName);
            measurentElement.AppendChild(measurentHibrid);

            xmlElements.Add(measurentElement);




            //Alt Sınır
            XmlElement low_limit_Element = xml.CreateElement("dcc", "quantity", dcc);
            low_limit_Element.SetAttribute("refType", ARFP_str + "_" + Coltext3);

            XmlElement low_limit_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement low_limit_Content = xml.CreateElement("dcc", "content", dcc);
            low_limit_Content.InnerText = ARFP_str + "_" + Coltext3;

            XmlElement low_limit_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement low_limit_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement low_limit_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement low_limit_Unit = xml.CreateElement("si", "unitXMLList", si);
            low_limit_Unit.InnerText = "\\dB";

            string low_limit_Data = string.Join(" ", col3.ToArray());
            low_limit_Value.InnerText = low_limit_Data;

            low_limit_Name.AppendChild(low_limit_Content);
            low_limit_RealList.AppendChild(low_limit_Value);
            low_limit_RealList.AppendChild(low_limit_Unit);
            low_limit_Hibrid.AppendChild(low_limit_RealList);
            low_limit_Element.AppendChild(low_limit_Name);
            low_limit_Element.AppendChild(low_limit_Hibrid);

            xmlElements.Add(low_limit_Element);


            // Sapma Fark Zayıflatma
            XmlElement SFZ_Element = xml.CreateElement("dcc", "quantity", dcc);
            SFZ_Element.SetAttribute("refType", ARFP_str + "_" + Coltext4);

            XmlElement SFZ_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement SFZ_Content = xml.CreateElement("dcc", "content", dcc);
            SFZ_Content.InnerText = ARFP_str + "_" + Coltext4;

            XmlElement SFZ_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement SFZ_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement SFZ_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement SFZ_Unit = xml.CreateElement("si", "unitXMLList", si);
            SFZ_Unit.InnerText = "\\dB";

            string SFZ_Data = string.Join(" ", col4.ToArray());
            SFZ_Value.InnerText = SFZ_Data;

            SFZ_Name.AppendChild(SFZ_Content);
            SFZ_RealList.AppendChild(SFZ_Value);
            SFZ_RealList.AppendChild(SFZ_Unit);
            SFZ_Hibrid.AppendChild(SFZ_RealList);
            SFZ_Element.AppendChild(SFZ_Name);
            SFZ_Element.AppendChild(SFZ_Hibrid);

            xmlElements.Add(SFZ_Element);



            // Üst sınır 
            XmlElement Up_Limit_Element = xml.CreateElement("dcc", "quantity", dcc);
            Up_Limit_Element.SetAttribute("refType", ARFP_str + "_" + Coltext5);

            XmlElement Up_Limit_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement Up_Limit_Content = xml.CreateElement("dcc", "content", dcc);
            Up_Limit_Content.InnerText = ARFP_str + "_" + Coltext5;

            XmlElement Up_Limit_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement Up_Limit_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement Up_Limit_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement Up_Limit_Unit = xml.CreateElement("si", "unitXMLList", si);
            Up_Limit_Unit.InnerText = "\\dB";

            string Up_Limit_Data = string.Join(" ", col5.ToArray());
            Up_Limit_Value.InnerText = Up_Limit_Data;

            Up_Limit_Name.AppendChild(Up_Limit_Content);
            Up_Limit_RealList.AppendChild(Up_Limit_Value);
            Up_Limit_RealList.AppendChild(Up_Limit_Unit);
            Up_Limit_Hibrid.AppendChild(Up_Limit_RealList);
            Up_Limit_Element.AppendChild(Up_Limit_Name);
            Up_Limit_Element.AppendChild(Up_Limit_Hibrid);

            xmlElements.Add(Up_Limit_Element);


            // Belirsizlik
            XmlElement UncertaintyElement = xml.CreateElement("dcc", "quantity", dcc);
            UncertaintyElement.SetAttribute("refType", ARFP_str + "_" + Coltext6);

            XmlElement UncertaintyName = xml.CreateElement("dcc", "name", dcc);
            XmlElement UncertaintyContent = xml.CreateElement("dcc", "content", dcc);
            UncertaintyContent.InnerText = ARFP_str + "_" + Coltext6;

            XmlElement UncertaintyHibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement UncertaintyRealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement UncertaintyValue = xml.CreateElement("si", "valueXMLList", si);
            XmlElement UncertaintyUnit = xml.CreateElement("si", "unitXMLList", si);
            UncertaintyUnit.InnerText = "\\dB";

            string UncertaintyData = string.Join(" ", col6.ToArray());
            UncertaintyValue.InnerText = UncertaintyData;

            UncertaintyName.AppendChild(UncertaintyContent);
            UncertaintyRealList.AppendChild(UncertaintyValue);
            UncertaintyRealList.AppendChild(UncertaintyUnit);
            UncertaintyHibrid.AppendChild(UncertaintyRealList);
            UncertaintyElement.AppendChild(UncertaintyName);
            UncertaintyElement.AppendChild(UncertaintyHibrid);

            xmlElements.Add(UncertaintyElement);


            return xmlElements;

        }

        public List<XmlElement> Add_ARFP_2(XmlDocument xml, ArrayList col1, ArrayList col2, ArrayList col3, ArrayList col4, string RRFP_str, string Coltext1, string Coltext2, string Coltext3, string Coltext4)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //seviye
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // lin Element Oluşturulması
            XmlElement Level_Element = xml.CreateElement("dcc", "quantity", dcc);
            Level_Element.SetAttribute("refType", RRFP_str + "_" + Coltext1);

            XmlElement Level_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement Level_Content = xml.CreateElement("dcc", "content", dcc);
            Level_Content.InnerText = RRFP_str + "_" + Coltext1;

            XmlElement Level_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement Level_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement Level_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement Level_Unit = xml.CreateElement("si", "unitXMLList", si);
            Level_Unit.InnerText = "\\dB";

            string Level_Data = string.Join(" ", col1.ToArray());
            Level_Value.InnerText = Level_Data;

            Level_Name.AppendChild(Level_Content);
            Level_RealList.AppendChild(Level_Value);
            Level_RealList.AppendChild(Level_Unit);
            Level_Hibrid.AppendChild(Level_RealList);
            Level_Element.AppendChild(Level_Name);
            Level_Element.AppendChild(Level_Hibrid);

            xmlElements.Add(Level_Element);



            // OlculenDeger
            XmlElement Measured_Val_Element = xml.CreateElement("dcc", "quantity", dcc);
            Measured_Val_Element.SetAttribute("refType", RRFP_str + "_" + Coltext2);

            XmlElement Measured_Val_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement Measured_Val_Content = xml.CreateElement("dcc", "content", dcc);
            Measured_Val_Content.InnerText = RRFP_str + "_" + Coltext2;

            XmlElement Measured_Val_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement Measured_Val_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement Measured_Val_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement Measured_Val_Unit = xml.CreateElement("si", "unitXMLList", si);
            Measured_Val_Unit.InnerText = "\\dB";

            string Measured_Val_Data = string.Join(" ", col2.ToArray());
            Measured_Val_Value.InnerText = Measured_Val_Data;

            Measured_Val_Name.AppendChild(Measured_Val_Content);
            Measured_Val_RealList.AppendChild(Measured_Val_Value);
            Measured_Val_RealList.AppendChild(Measured_Val_Unit);
            Measured_Val_Hibrid.AppendChild(Measured_Val_RealList);
            Measured_Val_Element.AppendChild(Measured_Val_Name);
            Measured_Val_Element.AppendChild(Measured_Val_Hibrid);

            xmlElements.Add(Measured_Val_Element);


            // Maksimum Değer
            XmlElement Max_Val_Element = xml.CreateElement("dcc", "quantity", dcc);
            Max_Val_Element.SetAttribute("refType", RRFP_str + "_" + Coltext3);

            XmlElement Max_Val_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement Max_Val_Content = xml.CreateElement("dcc", "content", dcc);
            Max_Val_Content.InnerText = RRFP_str + "_" + Coltext3;

            XmlElement Max_Val_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement Max_Val_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement Max_Val_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement Max_Val_Unit = xml.CreateElement("si", "unitXMLList", si);
            Max_Val_Unit.InnerText = "\\dB";

            string Max_Val_Data = string.Join(" ", col3.ToArray());
            Max_Val_Value.InnerText = Max_Val_Data;

            Max_Val_Name.AppendChild(Max_Val_Content);
            Max_Val_RealList.AppendChild(Max_Val_Value);
            Max_Val_RealList.AppendChild(Max_Val_Unit);
            Max_Val_Hibrid.AppendChild(Max_Val_RealList);
            Max_Val_Element.AppendChild(Max_Val_Name);
            Max_Val_Element.AppendChild(Max_Val_Hibrid);

            xmlElements.Add(Max_Val_Element);


            // Belirsizlik
            XmlElement Uncertainty_Element = xml.CreateElement("dcc", "quantity", dcc);
            Uncertainty_Element.SetAttribute("refType", RRFP_str + "_" + Coltext4);

            XmlElement Uncertainty_Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement Uncertainty_Content = xml.CreateElement("dcc", "content", dcc);
            Uncertainty_Content.InnerText = RRFP_str + "_" + Coltext4;

            XmlElement Uncertainty_Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement Uncertainty_RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement Uncertainty_Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement Uncertainty_Unit = xml.CreateElement("si", "unitXMLList", si);
            Uncertainty_Unit.InnerText = "\\dB";

            string Uncertainty_Data = string.Join(" ", col4.ToArray());
            Uncertainty_Value.InnerText = Uncertainty_Data;

            Uncertainty_Name.AppendChild(Uncertainty_Content);
            Uncertainty_RealList.AppendChild(Uncertainty_Value);
            Uncertainty_RealList.AppendChild(Uncertainty_Unit);
            Uncertainty_Hibrid.AppendChild(Uncertainty_RealList);
            Uncertainty_Element.AppendChild(Uncertainty_Name);
            Uncertainty_Element.AppendChild(Uncertainty_Hibrid);

            xmlElements.Add(Uncertainty_Element);

            return xmlElements;

        }
        #endregion

        #region RF DİFFERENCE
        public XmlDocument Add_RFD_result(XmlDocument xml, string str, XML_Arrays dataXml, List<bool> control)
        {
            //Result Namespace oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", dcc);
            nsmgr.AddNamespace("si", si);

            this.dataList = control;

            XmlNode sResults = xml.SelectSingleNode("/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results", nsmgr);

            // Elementlerin oluşturulması
            XmlElement result = xml.CreateElement("dcc", "result", dcc);
            result.SetAttribute("id", "RF_Difference" + str + "_dB");
            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement data = xml.CreateElement("dcc", "data", dcc);
            XmlElement list = xml.CreateElement("dcc", "list", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "RF_Difference" + str + "_dB";

            //Dataları içeren elementler oluşturulup dcc:List elementine eklenir.


            if (control[0])
            {
                list.AppendChild(add_RFD_Frekans(xml, dataXml.XML_RFD_T1_Frekans, "MaxOutPowTest"));

                List<XmlElement> xmlList = Add_RF_Difference(xml, dataXml.XML_RFD_T1_GostergeDegeri, dataXml.XML_RFD_T1_AltSınır, dataXml.XML_RFD_T1_OlculenDeger, dataXml.XML_RFD_T1_OlculenFark, dataXml.XML_RFD_T1_ÜstSınır, dataXml.XML_RFD_T1_Belirsizlik,
                                           "RFD", "IndıcatorVal", "lowerLimit", "measuredVal", "measureDiff", "upperLimit", "uncertainty");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }

            if (control[1])
            {
                list.AppendChild(add_RFD_Frekans(xml, dataXml.XML_RFD_T2_Frekans, "LevelAccTestFreq"));

                List<XmlElement> xmlList = Add_RF_Difference(xml, dataXml.XML_RFD_T2_Nom_Guc_Lvl, dataXml.XML_RFD_T2_OlculenDeger, dataXml.XML_RFD_T2_AltSınır, dataXml.XML_RFD_T2_Nom_Guc_Lvl_fark, dataXml.XML_RFD_T2_ÜstSınır, dataXml.XML_RFD_T2_Belirsizlik,
                                           "RFD", "NomPowlvl", "measuredVal", "lowerLimit", "NomPowlvlDiff", "upperLimit", "uncertainty");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[2])
            {
                list.AppendChild(add_RFD_Frekans(xml, dataXml.XML_RFD_T3_Frekans, "LevelAccTestPowRange"));

                List<XmlElement> xmlList = Add_RF_Difference(xml, dataXml.XML_RFD_T3_NominalGuc, dataXml.XML_RFD_T3_AltSınır, dataXml.XML_RFD_T3_OlculenDeger, dataXml.XML_RFD_T3_ÜstSınır, dataXml.XML_RFD_T3_Fark, dataXml.XML_RFD_T3_Belirsizlik,
                                           "RFD", "Nom_pow", "lowerLimit", "measuredVal", "upperLimit", "difference", "uncertainty");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[3])
            {
                list.AppendChild(add_RFD_Frekans(xml, dataXml.XML_RFD_T4_Frekans, "Summary"));

                List<XmlElement> xmlList = Add_RF_Difference(xml, dataXml.XML_RFD_T4_Min_Guc_lvl, dataXml.XML_RFD_T4_Max_Guc_lvl, dataXml.XML_RFD_T4_AltSınır, dataXml.XML_RFD_T4_Fark, dataXml.XML_RFD_T4_UstSınır, dataXml.XML_RFD_T4_Belirsizlik,
                                           "RFD", "MinPowLevel", "MaxPowLevel", "LowerLimit", "difference", "upper_limit", "uncertainty");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }



            //Son eklemeler yapılarak data geçirme tamamlanır.
            name.AppendChild(content);
            result.AppendChild(name);
            data.AppendChild(list);
            result.AppendChild(data);
            sResults.AppendChild(result);

            return xml;
        }


        public XmlElement add_RFD_Frekans(XmlDocument xml, ArrayList ArrayFrekans, string reftype)
        {
            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            XmlElement frekansElement = xml.CreateElement("dcc", "quantity", dcc);
            frekansElement.SetAttribute("refType", reftype);

            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "frequency";

            XmlElement hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement realList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement unit = xml.CreateElement("si", "unitXMLList", si);
            unit.InnerText = "\\dB";

            string frekansData = string.Join(" ", ArrayFrekans.ToArray());
            value.InnerText = frekansData;

            name.AppendChild(content);
            realList.AppendChild(value);
            realList.AppendChild(unit);
            hibrid.AppendChild(realList);
            frekansElement.AppendChild(name);
            frekansElement.AppendChild(hibrid);

            return frekansElement;

        }





        public List<XmlElement> Add_RF_Difference(XmlDocument xml, ArrayList COL1, ArrayList COL2, ArrayList COL3, ArrayList COL4, ArrayList COL5, ArrayList COL6, string RFD_str, string col1, string col2, string col3, string col4, string col5, string col6)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            // 1. SÜTUN
            XmlElement col1Element = xml.CreateElement("dcc", "quantity", dcc);
            col1Element.SetAttribute("refType", RFD_str + "_" + col1);

            XmlElement col1Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement col1Content = xml.CreateElement("dcc", "content", dcc);
            col1Content.InnerText = RFD_str + "_" + col1;

            XmlElement col1Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement col1RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement col1Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement col1Unit = xml.CreateElement("si", "unitXMLList", si);
            col1Unit.InnerText = "\\dB";

            string col1Data = string.Join(" ", COL1.ToArray());
            col1Value.InnerText = col1Data;

            col1Name.AppendChild(col1Content);
            col1RealList.AppendChild(col1Value);
            col1RealList.AppendChild(col1Unit);
            col1Hibrid.AppendChild(col1RealList);
            col1Element.AppendChild(col1Name);
            col1Element.AppendChild(col1Hibrid);

            xmlElements.Add(col1Element);

            // 2. SÜTUN
            XmlElement col2Element = xml.CreateElement("dcc", "quantity", dcc);
            col2Element.SetAttribute("refType", RFD_str + "_" + col2);

            XmlElement col2Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement col2Content = xml.CreateElement("dcc", "content", dcc);
            col2Content.InnerText = RFD_str + "_" + col2;

            XmlElement col2Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement col2RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement col2Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement col2Unit = xml.CreateElement("si", "unitXMLList", si);
            col2Unit.InnerText = "\\dB";

            string col2Data = string.Join(" ", COL2.ToArray());
            col2Value.InnerText = col2Data;

            col2Name.AppendChild(col2Content);
            col2RealList.AppendChild(col2Value);
            col2RealList.AppendChild(col2Unit);
            col2Hibrid.AppendChild(col2RealList);
            col2Element.AppendChild(col2Name);
            col2Element.AppendChild(col2Hibrid);

            xmlElements.Add(col2Element);

            // 3.SÜTUN
            XmlElement col3Element = xml.CreateElement("dcc", "quantity", dcc);
            col3Element.SetAttribute("refType", RFD_str + "_" + col3);

            XmlElement col3Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement col3Content = xml.CreateElement("dcc", "content", dcc);
            col3Content.InnerText = RFD_str + "_" + col3;

            XmlElement col3Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement col3RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement col3Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement col3Unit = xml.CreateElement("si", "unitXMLList", si);
            col3Unit.InnerText = "\\dB";

            string col3Data = string.Join(" ", COL3.ToArray());
            col3Value.InnerText = col3Data;

            col3Name.AppendChild(col3Content);
            col3RealList.AppendChild(col3Value);
            col3RealList.AppendChild(col3Unit);
            col3Hibrid.AppendChild(col3RealList);
            col3Element.AppendChild(col3Name);
            col3Element.AppendChild(col3Hibrid);

            xmlElements.Add(col3Element);

            // 4. SÜTUN
            XmlElement col4Element = xml.CreateElement("dcc", "quantity", dcc);
            col4Element.SetAttribute("refType", RFD_str + "_" + col4);

            XmlElement col4Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement col4Content = xml.CreateElement("dcc", "content", dcc);
            col4Content.InnerText = RFD_str + "_" + col4;

            XmlElement col4Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement col4RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement col4Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement col4Unit = xml.CreateElement("si", "unitXMLList", si);
            col4Unit.InnerText = "\\dB";

            string col4Data = string.Join(" ", COL4.ToArray());
            col4Value.InnerText = col4Data;

            col4Name.AppendChild(col4Content);
            col4RealList.AppendChild(col4Value);
            col4RealList.AppendChild(col4Unit);
            col4Hibrid.AppendChild(col4RealList);
            col4Element.AppendChild(col4Name);
            col4Element.AppendChild(col4Hibrid);

            xmlElements.Add(col4Element);

            // 5.SÜTUN
            XmlElement col5Element = xml.CreateElement("dcc", "quantity", dcc);
            col5Element.SetAttribute("refType", RFD_str + "_" + col5);

            XmlElement col5Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement col5Content = xml.CreateElement("dcc", "content", dcc);
            col5Content.InnerText = RFD_str + "_" + col5;

            XmlElement col5Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement col5RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement col5Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement col5Unit = xml.CreateElement("si", "unitXMLList", si);
            col5Unit.InnerText = "\\dB";

            string col5Data = string.Join(" ", COL5.ToArray());
            col5Value.InnerText = col5Data;

            col5Name.AppendChild(col5Content);
            col5RealList.AppendChild(col5Value);
            col5RealList.AppendChild(col5Unit);
            col5Hibrid.AppendChild(col5RealList);
            col5Element.AppendChild(col5Name);
            col5Element.AppendChild(col5Hibrid);

            xmlElements.Add(col5Element);





            // 6.SÜTUN
            XmlElement col6Element = xml.CreateElement("dcc", "quantity", dcc);
            col6Element.SetAttribute("refType", RFD_str + "_" + col6);

            XmlElement col6Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement col6Content = xml.CreateElement("dcc", "content", dcc);
            col6Content.InnerText = RFD_str + "_" + col6;

            XmlElement col6Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement col6RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement col6Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement col6Unit = xml.CreateElement("si", "unitXMLList", si);
            col6Unit.InnerText = "\\dB";

            string col6Data = string.Join(" ", COL6.ToArray());
            col6Value.InnerText = col6Data;

            col6Name.AppendChild(col6Content);
            col6RealList.AppendChild(col6Value);
            col6RealList.AppendChild(col6Unit);
            col6Hibrid.AppendChild(col6RealList);
            col6Element.AppendChild(col6Name);
            col6Element.AppendChild(col6Hibrid);

            xmlElements.Add(col6Element);



            return xmlElements;

        }
        #endregion

        #region RF GAİN
        public XmlDocument Add_RFG_result(XmlDocument xml, string str, XML_Arrays dataXml, List<bool> control)
        {
            //Result Namespace oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", dcc);
            nsmgr.AddNamespace("si", si);

            this.dataList = control;

            XmlNode sResults = xml.SelectSingleNode("/dcc:digitalCalibrationCertificate/dcc:measurementResults/dcc:measurementResult/dcc:results", nsmgr);

            // Elementlerin oluşturulması
            XmlElement result = xml.CreateElement("dcc", "result", dcc);
            result.SetAttribute("id", "RF_Gain" + str + "_dB");
            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement data = xml.CreateElement("dcc", "data", dcc);
            XmlElement list = xml.CreateElement("dcc", "list", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = "RF_Gain" + str + "_dB";

            //Dataları içeren elementler oluşturulup dcc:List elementine eklenir.


            if (control[0])
            {
                list.AppendChild(add_RFG_Frekans(xml, dataXml.XML_RFG_T1_Frekans, "Gain_input_nom_freq"));

                List<XmlElement> xmlList = Add_RF_Gain(xml, dataXml.XML_RFG_T1_GirisGucu, dataXml.XML_RFG_T1_Belirsizlik, "RFG", "Input_Pow1", "Unc1");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }

            if (control[1])
            {
                list.AppendChild(add_RFG_Frekans(xml, dataXml.XML_RFG_T2_EnBuyukKazanc, "Biggest_gain"));

                List<XmlElement> xmlList = Add_RF_Gain(xml, dataXml.XML_RFG_T2_EnKucukKazanc, dataXml.XML_RFG_T2_Flatness,
                                           "RFG", "lowest_Gain", "Flatness");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[2])
            {
                list.AppendChild(add_RFG_Frekans(xml, dataXml.XML_RFG_T3_Nom_Giris_Gucu, "Gain_diff_input_100KHz"));

                List<XmlElement> xmlList = Add_RF_Gain(xml, dataXml.XML_RFG_T3_Kazanc, dataXml.XML_RFG_T3_Belirsizlik, "RFG", "Input_Pow2", "Unc2");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }
            if (control[3])
            {
                list.AppendChild(add_RFG_Frekans(xml, dataXml.XML_RFG_T4_Nom_Giris_Gucu, "Gain_diff_input_1GHz"));

                List<XmlElement> xmlList = Add_RF_Gain(xml, dataXml.XML_RFG_T4_Kazanc, dataXml.XML_RFG_T4_Belirsizlik, "RFG", "Input_Pow3", "Unc3");
                foreach (XmlElement xmlElement in xmlList)
                {
                    list.AppendChild(xmlElement);
                }
            }



            //Son eklemeler yapılarak data geçirme tamamlanır.
            name.AppendChild(content);
            result.AppendChild(name);
            data.AppendChild(list);
            result.AppendChild(data);
            sResults.AppendChild(result);

            return xml;
        }


        public XmlElement add_RFG_Frekans(XmlDocument xml, ArrayList ArrayFrekans, string reftype)
        {
            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");

            // Node oluşturulması
            XmlElement frekansElement = xml.CreateElement("dcc", "quantity", dcc);
            frekansElement.SetAttribute("refType", reftype);

            XmlElement name = xml.CreateElement("dcc", "name", dcc);
            XmlElement content = xml.CreateElement("dcc", "content", dcc);
            content.InnerText = reftype;

            XmlElement hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement realList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement unit = xml.CreateElement("si", "unitXMLList", si);
            unit.InnerText = "\\dB";

            string frekansData = string.Join(" ", ArrayFrekans.ToArray());
            value.InnerText = frekansData;

            name.AppendChild(content);
            realList.AppendChild(value);
            realList.AppendChild(unit);
            hibrid.AppendChild(realList);
            frekansElement.AppendChild(name);
            frekansElement.AppendChild(hibrid);

            return frekansElement;

        }





        public List<XmlElement> Add_RF_Gain(XmlDocument xml, ArrayList COL1, ArrayList COL2, string RFG_str, string col1, string col2)
        {
            List<XmlElement> xmlElements = new List<XmlElement>();

            //Namespace manager oluşturma
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("dcc", "https://ptb.de/dcc");
            nsmgr.AddNamespace("si", "https://ptb.de/si");


            // 1. SÜTUN
            XmlElement col1Element = xml.CreateElement("dcc", "quantity", dcc);
            col1Element.SetAttribute("refType", RFG_str + "_" + col1);

            XmlElement col1Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement col1Content = xml.CreateElement("dcc", "content", dcc);
            col1Content.InnerText = RFG_str + "_" + col1;

            XmlElement col1Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement col1RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement col1Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement col1Unit = xml.CreateElement("si", "unitXMLList", si);
            col1Unit.InnerText = "\\dB";

            string col1Data = string.Join(" ", COL1.ToArray());
            col1Value.InnerText = col1Data;

            col1Name.AppendChild(col1Content);
            col1RealList.AppendChild(col1Value);
            col1RealList.AppendChild(col1Unit);
            col1Hibrid.AppendChild(col1RealList);
            col1Element.AppendChild(col1Name);
            col1Element.AppendChild(col1Hibrid);

            xmlElements.Add(col1Element);

            // 2. SÜTUN
            XmlElement col2Element = xml.CreateElement("dcc", "quantity", dcc);
            col2Element.SetAttribute("refType", RFG_str + "_" + col2);

            XmlElement col2Name = xml.CreateElement("dcc", "name", dcc);
            XmlElement col2Content = xml.CreateElement("dcc", "content", dcc);
            col2Content.InnerText = RFG_str + "_" + col2;

            XmlElement col2Hibrid = xml.CreateElement("si", "hybrid", si);
            XmlElement col2RealList = xml.CreateElement("si", "realListXMLList", si);
            XmlElement col2Value = xml.CreateElement("si", "valueXMLList", si);
            XmlElement col2Unit = xml.CreateElement("si", "unitXMLList", si);
            col2Unit.InnerText = "\\dB";

            string col2Data = string.Join(" ", COL2.ToArray());
            col2Value.InnerText = col2Data;

            col2Name.AppendChild(col2Content);
            col2RealList.AppendChild(col2Value);
            col2RealList.AppendChild(col2Unit);
            col2Hibrid.AppendChild(col2RealList);
            col2Element.AppendChild(col2Name);
            col2Element.AppendChild(col2Hibrid);

            xmlElements.Add(col2Element);

            return xmlElements;

        }
        #endregion
    }
}


