﻿using System;
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
            phaseUncElement.SetAttribute("refType", "Effective Effiency " + EE + "Imaginer_Unc");

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
            cfrelContent.InnerText = cfstr ;

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
            CIS_ZP_Unc_Element.SetAttribute("refType", "Calculable Impedance Standard " + CIS + "-Z-Position Unc");

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
            OCID_Unc_Element.SetAttribute("refType", "Calculable Impedance Standard  " + CIS + "-OCID Unc");

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
    }
}


