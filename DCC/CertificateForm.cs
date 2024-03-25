

using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DCC;
using System.Xml;

namespace DCC
{

    public partial class CertificateForm : Form
    {
        #region Tanımlamalar

        OpenFileDialog openFileDialog = new OpenFileDialog();
        public string ExcelDosyaYolu;
        public string ExcelDosyaAdi;
        public string WordFolderPath;
        string pageName;
        public List<Table> tables = new List<Table>();
        public List<string> header = new List<string>();
        public string XMLFolderPath;
        public XmlDocument xml = new XmlDocument();
        SP_DataWord sp_DataWord = new SP_DataWord();
        string TableName;
        XML_Arrays XML_Arrays = new XML_Arrays();
        SP_WordTable SP_WordTable = new SP_WordTable();
        EE_DataWord EE_DataWord = new EE_DataWord();
        CreateXML CreateXML = new CreateXML();
        CF_DataWord CF_DataWord = new CF_DataWord();
        EE_WordTable EE_WordTable = new EE_WordTable();
        CIS_DataWord CIS_DataWord = new CIS_DataWord();
        CF_WordTable CF_Word_Table = new CF_WordTable();
        CIS_WordTable CIS_Word_Table = new CIS_WordTable();
        CreateTemplate createTemplate = new CreateTemplate();
        Noise_DataWord Noise_DataWord = new Noise_DataWord();
        Noise_WordTable Noise_WordTable = new Noise_WordTable();
        Absolute_RF_Power_DataWord Absolute_RF_Power = new Absolute_RF_Power_DataWord();
        AbsoluteRF_Power_Word_Table Absolute_WordTable = new AbsoluteRF_Power_Word_Table();
        int satır;
        string sütun;

        #endregion

        public CertificateForm(XmlDocument xml)
        {
            this.xml = xml;
            InitializeComponent();
        }
        #region API PAGE

        private void CertificateForm_Load(object sender, EventArgs e)
        {
            LaboratoryComboBox.Enabled = false;
            DeviceNameTextBox.Enabled = false;
            ModelNameTextBox.Enabled = false;
            SerialNumberTextBox.Enabled = false;
            CalCodeTextBox.Enabled = false;
            SelectDeviceButton.Enabled = false;
            DeviceTextBox.Enabled = false;
            MethodTextBox.Enabled = false;
            CalibrationDescTextBox.Enabled = false;
            MeasurementsTextBox.Enabled = false;
            ReceiveData_Button.Enabled = false;
            CreateCertificate_Button.Enabled = false;
            BackBox3.Enabled = true;


        }

        private void OrderNumberTextBox_TextChanged(object sender, EventArgs e)
        {
            LaboratoryComboBox.Enabled = true;
        }

        private void LaboratoryComboBox_TextChanged(object sender, EventArgs e)
        {
            DeviceNameTextBox.Enabled = true;
        }

        private void DeviceNameTextBox_TextChanged(object sender, EventArgs e)
        {
            ModelNameTextBox.Enabled = true;
        }

        private void ModelNameTextBox_TextChanged(object sender, EventArgs e)
        {
            SerialNumberTextBox.Enabled = true;
        }

        private void SerialNumberTextBox_TextChanged(object sender, EventArgs e)
        {
            CalCodeTextBox.Enabled = true;

        }

        private void CalCodeTextBox_TextChanged(object sender, EventArgs e)
        {
            SelectDeviceButton.Enabled = true;
        }

        private void SelectDeviceButton_Click(object sender, EventArgs e)
        {
            DeviceTextBox.Enabled = true;
        }

        private void DeviceTextBox_TextChanged(object sender, EventArgs e)
        {
            MethodTextBox.Enabled = true;
        }

        private void MethodTextBox_TextChanged(object sender, EventArgs e)
        {
            CalibrationDescTextBox.Enabled = true;
        }

        private void CalibrationDescTextBox_TextChanged(object sender, EventArgs e)
        {
            MeasurementsTextBox.Enabled = true;
        }
        #endregion

        #region Next,Back Button ve Combobox Kontrolleri 
        private void NextButton_Click(object sender, EventArgs e)
        {
            this.Text = "DATA OPERATIONS";
            CertificateTabControl.SelectedTab = DATA_PAGE;
        }

        private void BackBox1_Click(object sender, EventArgs e)
        {
            HomePage homePage = new HomePage();
            this.Visible = false;
            homePage.Visible = true;
            homePage.Show();
        }

        private void BackBox2_Click(object sender, EventArgs e)
        {
            CertificateTabControl.SelectedTab = API_PAGE;
            this.Text = "API OPERATIONS";
        }

        private void BackBox3_Click(object sender, EventArgs e)
        {
            this.Text = "DATA OPERATIONS";
            CertificateTabControl.SelectedTab = DATA_PAGE;
            label4.Text = "Please double click on the cell to select it.";
            label4.Location = new System.Drawing.Point(377, 53);
        }

        private void NextBox1_Click(object sender, EventArgs e)
        {
            CertificateTabControl.SelectedTab = ExcelView_Page;
        }

        private void MeasurementTypes_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (MeasurementTypes_ComboBox.SelectedIndex == 0)
            {
                CheckBoxTabControl.SelectedTab = EE_Page;
            }
            if (MeasurementTypes_ComboBox.SelectedIndex == 1)
            {
                CheckBoxTabControl.SelectedTab = CalFactor_Page;
            }
            if (MeasurementTypes_ComboBox.SelectedIndex == 2)
            {
                CheckBoxTabControl.SelectedTab = CIS_Page;
            }
            if (MeasurementTypes_ComboBox.SelectedIndex == 3)
            {
                CheckBoxTabControl.SelectedTab = Absolute_RFPow_Page;
            }
            if (MeasurementTypes_ComboBox.SelectedIndex == 4)
            {
                CheckBoxTabControl.SelectedTab = RF_Difference_Tabpage;
            }        
            if (MeasurementTypes_ComboBox.SelectedIndex == 5)
            {
                CheckBoxTabControl.SelectedTab = RF_Gain_Tabpage;
            }
            if (MeasurementTypes_ComboBox.SelectedIndex == 6)
            {
                CheckBoxTabControl.SelectedTab = SParam_Page;
            }
            if (MeasurementTypes_ComboBox.SelectedIndex == 7)
            {
                CheckBoxTabControl.SelectedTab = Noise_Page;
            }
            if (MeasurementTypes_ComboBox.SelectedIndex == 8)
            {
                CheckBoxTabControl.SelectedTab = MetCH_Page;
            }
        }
        #endregion


        private void ExcelPage_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Text = "EXCEL ROW AND COLUMN SELECT";
            CertificateTabControl.SelectedTab = ExcelView_Page;

            string selectedWorksheetName = ExcelPage_ComboBox.SelectedItem as string;
            if (selectedWorksheetName != null)
            {
                DisplayExcelWorksheet(selectedWorksheetName);
            }
            pageName = ExcelPage_ComboBox.SelectedItem.ToString();
            CertificateTabControl.SelectedTab = ExcelView_Page;

            ReceiveData_Button.Enabled = true;
        }
        #region Button
        private void Save_Row_Col_Button_Click(object sender, EventArgs e)
        {
            if (satır == 0 && sütun == null)
            {
                label4.Location = new System.Drawing.Point(475, 53);

                label4.Text = "Please select a cell";
            }
            else
            {
                sütun = sütun.ToUpper();
                this.Text = "DATA OPERATIONS";
                CertificateTabControl.SelectedTab = DATA_PAGE;
                label4.Text = "Please double click on the cell to select it.";
            }
        }

        private void SelectExcel_Button_Click(object sender, EventArgs e)
        {
            {
                LabelProgress.Visible = false;
                progressBar.Value = 0;
                try
                {


                    openFileDialog.Filter = @"Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                    openFileDialog.InitialDirectory = "C:\\";
                    openFileDialog.Title = @"Excel Files, Select a "".xlsx"" file";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExcelDosyaYolu = openFileDialog.FileName;
                        ExcelDosyaAdi = Path.GetFileName(ExcelDosyaYolu);
                        ExcelFileName_TextBox.Text = ExcelDosyaAdi;
                        // Excel dosyasını oku ve sayfa isimlerini ComboBox'a ekle
                        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        using (var excelPackage = new OfficeOpenXml.ExcelPackage(new FileInfo(ExcelDosyaYolu)))
                        {
                            // ComboBox'ı temizle
                            ExcelPage_ComboBox.Items.Clear();

                            // Excel dosyasındaki tüm sayfa isimlerini ComboBox'a ekle
                            foreach (var sayfa in excelPackage.Workbook.Worksheets)
                            {
                                ExcelPage_ComboBox.Items.Add(sayfa.Name);
                            }

                        }

                        // Feedback işlemi
                        LabelProgress.Visible = true;
                        LabelProgress.ForeColor = System.Drawing.Color.Green;
                        LabelProgress.Text = @"Excel file selection successful";
                    }
                }
                catch (Exception err)
                {
                    LabelProgress.Visible = true;
                    LabelProgress.ForeColor = System.Drawing.Color.Red;
                    LabelProgress.Text = @"ERROR!: Selection of Excel";
                    MessageBox.Show(err.Message, err.StackTrace, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }



            }
        }

        private void ReceiveData_Button_Click(object sender, EventArgs e)
        {
            CreateCertificate_Button.Enabled = true;

            #region dataword çalıştırma
            try
            {
                #region S-Parametre
                if (MeasurementTypes_ComboBox.SelectedIndex == 6)
                {
                    sp_DataWord.main(ExcelDosyaYolu, pageName, satır, sütun);
                    XML_Arrays.SP_Data_Xml(ExcelDosyaYolu, pageName, satır, sütun);
                    listBox1.Items.Add("S Parametre-" + ExcelDosyaAdi);

                    #region S parametre Checkbox Kontrolleri

                    if (ExcelFileName_TextBox.Text != "(Please enter a header name..)")
                    {
                        TableName = MeasurementTypes_ComboBox.Text + " - ";
                    }
                    else
                    {
                        TableName = "";
                    }

                    List<bool> dataList = new List<bool>(14) { false, false, false, false, false, false, false, false, false, false, false, false, false, false };

                    if (checkBoxS11Reel.Checked)
                    {          // Reel & Imaginer kutusu
                        string txtS11Reel = TableName + "Reel and Imaginary Components for S11\n\n";
                        Table s11reelTable = SP_WordTable.CreateReelImg(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS11Reel, sp_DataWord.ArrayS11ReelUnc, sp_DataWord.ArrayS11Complex, sp_DataWord.ArrayS11ComplexUnc);
                        tables.Add(s11reelTable);
                        header.Add(txtS11Reel);
                        dataList[0] = true;
                        SaveBasarim();
                    }

                    if (checkBoxS11Lin.Checked)
                    {    // Linear kutusu
                        string txtS11Lin = TableName + "Linear Magnitude and Phase Components for S11\n\n";
                        Table s11linTable = SP_WordTable.CreateLinPhase(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS11Lin, sp_DataWord.ArrayS11LinUnc, sp_DataWord.ArrayS11LinPhase, sp_DataWord.ArrayS11LinPhaseUnc);
                        tables.Add(s11linTable);
                        header.Add(txtS11Lin);
                        dataList[1] = true;
                        SaveBasarim();
                    }
                    if (checkBoxS11Log.Checked)
                    {    // Logarithmic Kutusu
                        string txtS11Log = TableName + "Logarithmic Magnitude and Phase Components for S11\n\n";
                        Table s11logTable = SP_WordTable.CreateLogPhase(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS11Log, sp_DataWord.ArrayS11LogUnc, sp_DataWord.ArrayS11LogPhase, sp_DataWord.ArrayS11LogPhaseUnc);
                        tables.Add(s11logTable);
                        header.Add(txtS11Log);
                        dataList[2] = true;
                        SaveBasarim();
                    }
                    if (checkBoxS11SWR.Checked)
                    {    // SWR Kutusu
                        string txtS11SWR = TableName + "SWR for S11";
                        Table s11swrTable = SP_WordTable.CreateSWR(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS11SWR, sp_DataWord.ArrayS11SWRUnc);
                        tables.Add(s11swrTable);
                        header.Add(txtS11SWR);
                        dataList[3] = true;
                        SaveBasarim();
                    }


                    if (checkBoxS12Reel.Checked)
                    {  // Reel & Imaginer kutusu
                        string txtS12Reel = TableName + "Reel and Imaginary Components for S12\n";
                        Table s12reelTable = SP_WordTable.CreateReelImg(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS12Reel, sp_DataWord.ArrayS12ReelUnc, sp_DataWord.ArrayS12Complex, sp_DataWord.ArrayS12ComplexUnc);
                        tables.Add(s12reelTable);
                        header.Add(txtS12Reel);
                        dataList[4] = true;
                        SaveBasarim();
                    }
                    if (checkBoxS12Lin.Checked)
                    {    // Linear kutusu
                        string txtS12Lin = TableName + "Linear Magnitude and Phase Components for S12\n\n";
                        Table s12linTable = SP_WordTable.CreateLinPhase(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS12Lin, sp_DataWord.ArrayS12LinUnc, sp_DataWord.ArrayS12LinPhase, sp_DataWord.ArrayS12LinPhaseUnc);
                        tables.Add(s12linTable);
                        header.Add(txtS12Lin);
                        dataList[5] = true;
                        SaveBasarim();
                    }
                    if (checkBoxS12Log.Checked)
                    {    // Logarithmic Kutusu
                        string txtS12Log = TableName + "Logarithmic Magnitude and Phase Components for S12\n\n";
                        Table s12logTable = SP_WordTable.CreateLogPhase(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS12Log, sp_DataWord.ArrayS12LogUnc, sp_DataWord.ArrayS12LogPhase, sp_DataWord.ArrayS12LogPhaseUnc);
                        tables.Add(s12logTable);
                        header.Add(txtS12Log);
                        dataList[6] = true;
                        SaveBasarim();
                    }


                    if (checkBoxS21Reel.Checked)
                    {  // Reel & Imaginer kutusu
                        string txtS21Reel = TableName + "Reel and Imaginary Components for S21\n";
                        Table s21reelTable = SP_WordTable.CreateReelImg(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS21Reel, sp_DataWord.ArrayS21ReelUnc, sp_DataWord.ArrayS21Complex, sp_DataWord.ArrayS21ComplexUnc);
                        tables.Add(s21reelTable);
                        header.Add(txtS21Reel);
                        dataList[7] = true;
                        SaveBasarim();
                    }
                    if (checkBoxS21Lin.Checked)
                    {    // Linear kutusu
                        string txtS21Lin = TableName + "Linear Magnitude and Phase Components for S21\n\n";
                        Table s21linTable = SP_WordTable.CreateLinPhase(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS21Lin, sp_DataWord.ArrayS21LinUnc, sp_DataWord.ArrayS21LinPhase, sp_DataWord.ArrayS21LinPhaseUnc);
                        tables.Add(s21linTable);
                        header.Add(txtS21Lin);
                        dataList[8] = true;
                        SaveBasarim();
                    }
                    if (checkBoxS21Log.Checked)
                    {    // Logarithmic Kutusu
                        string txtS21Log = TableName + "Logarithmic Magnitude and Phase Components for S21\n\n";
                        Table s21logTable = SP_WordTable.CreateLogPhase(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS21Log, sp_DataWord.ArrayS21LogUnc, sp_DataWord.ArrayS21LogPhase, sp_DataWord.ArrayS21LogPhaseUnc);
                        tables.Add(s21logTable);
                        header.Add(txtS21Log);
                        dataList[9] = true;
                        SaveBasarim();
                    }

                    if (checkBoxS22Reel.Checked)
                    {  // Reel & Imaginer kutusu
                        string txtS22Reel = TableName + "Reel and Imaginary Components for S22\n";
                        Table s22reelTable = SP_WordTable.CreateReelImg(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS22Reel, sp_DataWord.ArrayS22ReelUnc, sp_DataWord.ArrayS22Complex, sp_DataWord.ArrayS22ComplexUnc);
                        tables.Add(s22reelTable);
                        header.Add(txtS22Reel);
                        dataList[10] = true;
                        SaveBasarim();
                    }
                    if (checkBoxS22Lin.Checked)
                    {    // Linear kutusu
                        string txtS22Lin = TableName + "Linear Magnitude and Phase Components for S22\n\n";
                        Table s22linTable = SP_WordTable.CreateLinPhase(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS22Lin, sp_DataWord.ArrayS22LinUnc, sp_DataWord.ArrayS22LinPhase, sp_DataWord.ArrayS22LinPhaseUnc);
                        tables.Add(s22linTable);
                        header.Add(txtS22Lin);
                        dataList[11] = true;
                        SaveBasarim();
                    }
                    if (checkBoxS22Log.Checked)
                    {    // Logarithmic Kutusu
                        string txtS22Log = TableName + "Logarithmic Magnitude and Phase Components for S22\n\n";
                        Table s22logTable = SP_WordTable.CreateLogPhase(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS22Log, sp_DataWord.ArrayS22LogUnc, sp_DataWord.ArrayS22LogPhase, sp_DataWord.ArrayS22LogPhaseUnc);
                        tables.Add(s22logTable);
                        header.Add(txtS22Log);
                        dataList[12] = true;
                        SaveBasarim();
                    }
                    if (checkBoxS22SWR.Checked)
                    {    // SWR Kutusu
                        string txtS22SWR = TableName + "SWR for S22";
                        Table s22swrTable = SP_WordTable.CreateSWR(sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS22SWR, sp_DataWord.ArrayS22SWRUnc);
                        tables.Add(s22swrTable);
                        header.Add(txtS22SWR);
                        dataList[13] = true;
                        SaveBasarim();
                    }

                    CreateXML.AddSParameterResult(xml, TableName, XML_Arrays, dataList);

                    #endregion


                }
                #endregion

                #region Effective Effiency

                else if (MeasurementTypes_ComboBox.SelectedIndex == 0)
                {
                    EE_DataWord.main(ExcelDosyaYolu, pageName, satır, sütun);
                    XML_Arrays.EE_Data_Xml(ExcelDosyaYolu, pageName, satır, sütun);
                    listBox1.Items.Add("Effective Effiency-" + ExcelDosyaAdi);

                    #region EE ChechkBox Kontrolleri
                    List<bool> dataListEE = new List<bool>(14) { false, false, false, false };


                    if (checkBoxEE.Checked)
                    {          // Reel & Imaginer kutusu
                        string txtS11Reel_EE = TableName + "S11 Reel And Imagıner Components for EE \n\n";
                        Table s11reelTable_EE = EE_WordTable.EECreateReelImg(EE_DataWord.EE_ArrayFrekans, EE_DataWord.EE_ArrayS11Reel, EE_DataWord.EE_ArrayS11ReelUnc, EE_DataWord.EE_ArrayS11Complex, EE_DataWord.EE_ArrayS11ComplexUnc);
                        tables.Add(s11reelTable_EE);
                        header.Add(txtS11Reel_EE);
                        dataListEE[0] = true;
                        SaveBasarim();
                    }

                    if (checkBox_EE_RI.Checked)
                    {    // Linear kutusu
                        string txt_EE = TableName + "Effective Effiency\n\n";
                        Table EE_table = EE_WordTable.CreateEE(EE_DataWord.EE_ArrayFrekans, EE_DataWord.EE_ArrayEE, EE_DataWord.EE_ArrayEEUnc);
                        tables.Add(EE_table);
                        header.Add(txt_EE);
                        dataListEE[1] = true;
                        SaveBasarim();
                    }

                    // Logarithmic Kutusu
                    if (checkBoxRHO.Checked)
                    {
                        string txtRHO = TableName + "RHO Tables\n\n";
                        Table RHOTable = EE_WordTable.CreateRHO(EE_DataWord.EE_ArrayFrekans, EE_DataWord.EE_ArrayRhoLin, EE_DataWord.EE_ArrayRhoUnc);
                        tables.Add(RHOTable);
                        header.Add(txtRHO);
                        dataListEE[2] = true;
                        SaveBasarim();
                    }
                    if (checkBox_EE_CF.Checked)
                    {    // SWR Kutusu
                        string txtCF_EE = TableName + "CF Tables for EE";
                        Table EE_CF_table = EE_WordTable.CreateCF(EE_DataWord.EE_ArrayFrekans, EE_DataWord.EE_ArrayCF, EE_DataWord.EE_ArrayCFUnc);
                        tables.Add(EE_CF_table);
                        header.Add(txtCF_EE);
                        dataListEE[3] = true;
                        SaveBasarim();
                    }
                    CreateXML.Add_EE_Result(xml, TableName, XML_Arrays, dataListEE);
                    #endregion

                }
                #endregion

                #region Calibration Factor
                else if (MeasurementTypes_ComboBox.SelectedIndex == 1)
                {
                    CF_DataWord.main(ExcelDosyaYolu, pageName, satır, sütun);
                    XML_Arrays.CF_Data_Xml(ExcelDosyaYolu, pageName, satır, sütun);
                    listBox1.Items.Add("Calibration Factor-" + ExcelDosyaAdi);
                    #region CF CheckBox Kontrolleri

                    List<bool> dataListCF = new List<bool>(14) { false, false };


                    if (CF_checkBox_RIRC.Checked)
                    {    // Linear kutusu
                        string txt_CF = TableName + "Reel, Imaginer and Reflection Coefficient\n";
                        Table CF_table = CF_Word_Table.CF_CreateCF(CF_DataWord.CF_ArrayFrekans, CF_DataWord.CF_Array, CF_DataWord.CF_ArrayCFUnc);
                        tables.Add(CF_table);
                        header.Add(txt_CF);
                        dataListCF[1] = true;
                        SaveBasarim();
                    }

                    if (CheckBox_CF.Checked)
                    {          // Reel & Imaginer kutusu
                        string txtReel_CF = TableName + "Calibration Factor \n\n";
                        Table reelTable_CF = CF_Word_Table.CF_CreateReelImg(CF_DataWord.CF_ArrayFrekans, CF_DataWord.CF_ArrayReel, CF_DataWord.CF_ArrayReelUnc, CF_DataWord.CF_ArrayComplex, CF_DataWord.CF_ArrayComplexUnc, CF_DataWord.CF_YK, CF_DataWord.CF_YK_Unc);
                        tables.Add(reelTable_CF);
                        header.Add(txtReel_CF);
                        dataListCF[0] = true;
                        SaveBasarim();
                    }

                    CreateXML.AddCFResult(xml, TableName, XML_Arrays, dataListCF);

                    #endregion
                }
                #endregion

                #region CIS
                else if (MeasurementTypes_ComboBox.SelectedIndex == 2)
                {

                    CIS_DataWord.main(ExcelDosyaYolu, pageName, satır, sütun);
                    XML_Arrays.CIS_Data_Xml(ExcelDosyaYolu, pageName, satır, sütun);
                    listBox1.Items.Add("Calculable Impedance Standard" + ExcelDosyaAdi);

                    bool CIS_bool = false;


                    if (CIS_CheckBox.Checked)
                    {    // Linear kutusu
                        string txt_CIS = TableName + "Z-Position ,OCID,ICOD\n";
                        Table CIS_table = CIS_Word_Table.Create_Z_Position(CIS_DataWord.CIS_Olcum_Adım, CIS_DataWord.CIS_ZP, CIS_DataWord.CIS_ZP_Unc, CIS_DataWord.CIS_ICOD, CIS_DataWord.CIS_ICOD_Unc, CIS_DataWord.CIS_OCID, CIS_DataWord.CIS_OCID_Unc);
                        tables.Add(CIS_table);
                        header.Add(txt_CIS);
                        CIS_bool = true;
                        SaveBasarim();
                    }
                    CreateXML.AddCISResult(xml, TableName, XML_Arrays, CIS_bool);

                }
                #endregion

                #region Noise
                else if (MeasurementTypes_ComboBox.SelectedIndex == 7)
                {
                    Noise_DataWord.main(ExcelDosyaYolu, pageName, satır, sütun);
                    XML_Arrays.Noise_Data_Xml(ExcelDosyaYolu, pageName, satır, sütun);
                    listBox1.Items.Add("Noise-" + ExcelDosyaAdi);


                    List<bool> NoiseBool = new List<bool>(3) { false, false, false };

                    if (NS_checkBoxENR.Checked)
                    {
                        string txt_ENR_Noise = TableName + "ENR, ENR Uncertainty\n";
                        Table ENR_Noise_table = Noise_WordTable.CreateENR(Noise_DataWord.NS_ArrayFrekans, Noise_DataWord.NS_ArrayENR, Noise_DataWord.NS_ArrayENRUnc);
                        tables.Add(ENR_Noise_table);
                        header.Add(txt_ENR_Noise);
                        NoiseBool[0] = true;
                        SaveBasarim();
                    }
                    if (NS_checkBox_DC_ON.Checked)
                    {
                        string txt_DC_ON_Noise = TableName + "DC ON for Noise\n";
                        Table DC_ON_Noise_table = Noise_WordTable.Create_DC_ON_OFF(Noise_DataWord.NS_ArrayFrekans, Noise_DataWord.NS_ArrayRC, Noise_DataWord.NS_ArrayRC_ustlimit, Noise_DataWord.NS_ArrayRCUnc,
                                                                                    Noise_DataWord.NS_ArrayRC_Phase, Noise_DataWord.NS_ArrayRC_PhaseUnc, Noise_DataWord.NS_ArrayControl_DC_ON);
                        tables.Add(DC_ON_Noise_table);
                        header.Add(txt_DC_ON_Noise);
                        NoiseBool[1] = true;
                        SaveBasarim();
                    }

                    if (NS_checkBox_DC_OFF.Checked)
                    {
                        string txt_DC_OFF_Noise = TableName + "DC OFF for Noise\n";
                        Table DC_OFF_Noise_table = Noise_WordTable.Create_DC_ON_OFF(Noise_DataWord.NS_ArrayFrekans, Noise_DataWord.NS_ArrayRC_DC_OFF, Noise_DataWord.NS_ArrayRC_ustlimit_DC_OFF, Noise_DataWord.NS_ArrayRCUnc_DC_OFF,
                                                                                    Noise_DataWord.NS_ArrayRC_Phase_DC_OFF, Noise_DataWord.NS_ArrayRC_PhaseUnc_DC_OFF, Noise_DataWord.NS_ArrayControl_DC_OFF);
                        tables.Add(DC_OFF_Noise_table);
                        header.Add(txt_DC_OFF_Noise);
                        NoiseBool[2] = true;
                        SaveBasarim();
                    }
                    CreateXML.AddNoiseResult(xml, TableName, XML_Arrays, NoiseBool);

                }
                #endregion

                #region Absolute RF Power
                else if (MeasurementTypes_ComboBox.SelectedIndex == 3)
                {
                    Absolute_RF_Power.main(ExcelDosyaYolu, pageName, satır, sütun);
                    XML_Arrays.ABS_RFP_Data_Xml(ExcelDosyaYolu, pageName, satır, sütun);

                    List<bool>  ARFPBool = new List<bool>(3) { false, false, false,false,false,false,false,false,false, false,false};

                    if (ARFP_1.Checked)
                    {
                        string txt_ARFP1 = TableName + "Head RF” Çıkışı Seviye Doğruluğu Testi\n";
                        Table ARFP1_table = Absolute_WordTable.ARFP_CreateTable_1(Absolute_RF_Power.ARFP_T1_Frekans, Absolute_RF_Power.ARFP_T1_Cıkıs_Gücü, Absolute_RF_Power.ARFP_T1_Olculen_Güc, Absolute_RF_Power.ARFP_T1_AltSınır,
                                                                                  Absolute_RF_Power.ARFP_T1_Sapma, Absolute_RF_Power.ARFP_T1_ÜstSınır, Absolute_RF_Power.ARFP_T1_Belirsizlik, "Ölçülen Güç (dBm)", "Sapma");
                        tables.Add(ARFP1_table);
                        header.Add(txt_ARFP1);
                        ARFPBool[0] = true;
                        SaveBasarim();
                    }

                    if (ARFP_2.Checked)
                    {
                        string txt_ARFP2 = TableName + "Güç Aralığına Göre Seviye Doğruluğu Testi\n";
                        Table ARFP2_table = Absolute_WordTable.ARFP_CreateTable_1(Absolute_RF_Power.ARFP_T2_Frekans, Absolute_RF_Power.ARFP_T2_Cıkıs_Gücü, Absolute_RF_Power.ARFP_T2_OlculenDeger, Absolute_RF_Power.ARFP_T2_AltSınır,
                                                                                  Absolute_RF_Power.ARFP_T2_Fark, Absolute_RF_Power.ARFP_T2_ÜstSınır, Absolute_RF_Power.ARFP_T2_Belirsizlik, "Ölçülen Değer (dBm)", "Fark (dB)");
                        tables.Add(ARFP2_table);
                        header.Add(txt_ARFP2);
                        ARFPBool[1] = true;
                        SaveBasarim();
                    }
                    if (ARFP_3.Checked)
                    {
                        string txt_ARFP3 = TableName + "Head RF” Çıkışı Zayıflatma Doğruluğu Testi\n";
                        Table ARFP3_table = Absolute_WordTable.ARFP_CreateTable_1(Absolute_RF_Power.ARFP_T3_Frekans, Absolute_RF_Power.ARFP_T3_Cıkıs_Gücü, Absolute_RF_Power.ARFP_T3_OlculenZayıflatma, Absolute_RF_Power.ARFP_T3_AltSınır,
                                                                                  Absolute_RF_Power.ARFP_T3_Zayıflatma, Absolute_RF_Power.ARFP_T3_ÜstSınır, Absolute_RF_Power.ARFP_T3_Belirsizlik, "Ölçülen Zayıflatma (dB)", "Zayıflatma Hatası (dB)");
                        tables.Add(ARFP3_table);
                        header.Add(txt_ARFP3);
                        ARFPBool[2] = true;
                        SaveBasarim();
                    }
                    if (ARFP_4.Checked)
                    {
                        string txt_ARFP4 = TableName + "Head RF” Çıkışı Duran Dalga Oranı (SWR) Testi @13dBm\n";
                        Table ARFP4_table = Absolute_WordTable.ARFP_CreateTable_2(Absolute_RF_Power.ARFP_T4_T5_T6_frekans, Absolute_RF_Power.ARFP_T4_SWR_Seviye, Absolute_RF_Power.ARFP_T4_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T4_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T4_SWR_Belirsizlik,"Maksimum Değer");
                        tables.Add(ARFP4_table);
                        header.Add(txt_ARFP4);
                        ARFPBool[3] = true;
                        SaveBasarim();
                    }
                    if (ARFP_5.Checked)
                    {
                        string txt_ARFP5 = TableName + "Head RF” Çıkışı Duran Dalga Oranı (SWR) Testi @3dBm\n";
                        Table ARFP5_table = Absolute_WordTable.ARFP_CreateTable_2(Absolute_RF_Power.ARFP_T4_T5_T6_frekans, Absolute_RF_Power.ARFP_T5_SWR_Seviye, Absolute_RF_Power.ARFP_T5_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T5_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T5_SWR_Belirsizlik, "Maksimum Değer");
                        tables.Add(ARFP5_table);
                        header.Add(txt_ARFP5);
                        ARFPBool[4] = true;
                        SaveBasarim();
                    }
                    if (ARFP_6.Checked)
                    {
                        string txt_ARFP6 = TableName + "Head RF” Çıkışı Duran Dalga Oranı (SWR) Testi @-7dBm\n";
                        Table ARFP6_table = Absolute_WordTable.ARFP_CreateTable_2(Absolute_RF_Power.ARFP_T4_T5_T6_frekans, Absolute_RF_Power.ARFP_T6_SWR_Seviye, Absolute_RF_Power.ARFP_T6_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T6_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T6_SWR_Belirsizlik, "Maksimum Değer");
                        tables.Add(ARFP6_table);
                        header.Add(txt_ARFP6);
                        ARFPBool[5] = true;
                        SaveBasarim();
                    }
                    if (ARFP_7.Checked)
                    {
                        string txt_ARFP7 = TableName + "Mikrodalga Çıkışı Seviye Doğruluğu Testi \n";
                        Table ARFP7_table = Absolute_WordTable.ARFP_CreateTable_1(Absolute_RF_Power.ARFP_T7_Frekans, Absolute_RF_Power.ARFP_T7_Cıkıs_Gücü, Absolute_RF_Power.ARFP_T7_OlculenGuc, Absolute_RF_Power.ARFP_T7_AltSınır,
                                                                                  Absolute_RF_Power.ARFP_T7_Sapma, Absolute_RF_Power.ARFP_T7_ÜstSınır, Absolute_RF_Power.ARFP_T7_Belirsizlik, "Ölçülen Güç (dBm)", "Sapma(dB)");
                        tables.Add(ARFP7_table);
                        header.Add(txt_ARFP7);
                        ARFPBool[6] = true;
                        SaveBasarim();
                    }
                    if (ARFP_8.Checked)
                    {
                        string txt_ARFP8 = TableName + "Güç Aralığına Göre Mikrodalga Çıkışı Seviye Doğruluğu Testi\n";
                        Table ARFP8_table = Absolute_WordTable.ARFP_CreateTable_1(Absolute_RF_Power.ARFP_T8_Frekans, Absolute_RF_Power.ARFP_T8_Cıkıs_Gücü, Absolute_RF_Power.ARFP_T8_OlculenDeger, Absolute_RF_Power.ARFP_T8_AltSınır,
                                                                                  Absolute_RF_Power.ARFP_T8_Fark, Absolute_RF_Power.ARFP_T8_ÜstSınır, Absolute_RF_Power.ARFP_T8_Belirsizlik, "Ölçülen Değer (dBm)", "Fark (dB)");
                        tables.Add(ARFP8_table);
                        header.Add(txt_ARFP8);
                        ARFPBool[7] = true;
                        SaveBasarim();
                    }
                    if (ARFP_9.Checked)
                    {
                        string txt_ARFP9 = TableName + "Mikrodalga Çıkışı Duran Dalga Oranı (SWR) Testi  @11 dBm\n";
                        Table ARFP9_table = Absolute_WordTable.ARFP_CreateTable_2(Absolute_RF_Power.ARFP_T9_T10_T11_frekans, Absolute_RF_Power.ARFP_T9_SWR_Seviye, Absolute_RF_Power.ARFP_T9_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T9_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T9_SWR_Belirsizlik, "Üst Sınır");
                        tables.Add(ARFP9_table);
                        header.Add(txt_ARFP9);
                        ARFPBool[8] = true;
                        SaveBasarim();
                    }
                    if (ARFP_10.Checked)
                    {
                        string txt_ARFP10 = TableName + "Mikrodalga Çıkışı Duran Dalga Oranı (SWR) Testi  @3 dBm\n";
                        Table ARFP10_table = Absolute_WordTable.ARFP_CreateTable_2(Absolute_RF_Power.ARFP_T9_T10_T11_frekans, Absolute_RF_Power.ARFP_T10_SWR_Seviye, Absolute_RF_Power.ARFP_T10_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T10_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T10_SWR_Belirsizlik, "Üst Sınır");
                        tables.Add(ARFP10_table);
                        header.Add(txt_ARFP10);
                        ARFPBool[9] = true;
                        SaveBasarim();
                    }
                    if (ARFP_11.Checked)
                    {
                        string txt_ARFP11 = TableName + "Mikrodalga Çıkışı Duran Dalga Oranı (SWR) Testi  @-9 dBm\n";
                        Table ARFP11_table = Absolute_WordTable.ARFP_CreateTable_2(Absolute_RF_Power.ARFP_T9_T10_T11_frekans, Absolute_RF_Power.ARFP_T11_SWR_Seviye, Absolute_RF_Power.ARFP_T11_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T11_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T11_SWR_Belirsizlik, "Üst Sınır");
                        tables.Add(ARFP11_table);
                        header.Add(txt_ARFP11);
                        ARFPBool[10] = true;
                        SaveBasarim();
                    }
                    CreateXML.Add_ARFP_Result(xml, TableName, XML_Arrays, ARFPBool);

                }

                #endregion

                    for (int i = 0; i < 100; i++)
                {

                }
                LabelProgress.Visible = true;
                LabelProgress.ForeColor = System.Drawing.Color.Green;
                LabelProgress.Text = @"Import data successfull";
            }

            catch (Exception err)
            {
                LabelProgress.Visible = true;
                LabelProgress.ForeColor = System.Drawing.Color.Red;
                LabelProgress.Text = @"ERROR!: Excel";
                MessageBox.Show(err.Message, err.StackTrace, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            #endregion

            #region Refresh

            DialogResult result = MessageBox.Show("Information have been saved.\nIf you want to add more results click Yes.\nIf not click No.", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            checkBoxS11Reel.Checked = false; checkBoxS12Reel.Checked = false; checkBoxS21Reel.Checked = false; checkBoxS22Reel.Checked = false;
            checkBoxS11Lin.Checked = false; checkBoxS12Lin.Checked = false; checkBoxS21Lin.Checked = false; checkBoxS22Lin.Checked = false;
            checkBoxS11Log.Checked = false; checkBoxS12Log.Checked = false; checkBoxS21Log.Checked = false; checkBoxS22Log.Checked = false;
            checkBoxS11SWR.Checked = false; checkBoxS22SWR.Checked = false;

            checkBoxEE.Checked = false; checkBox_EE_RI.Checked = false; checkBoxRHO.Checked = false; checkBox_EE_CF.Checked = false;
            CF_checkBox_RIRC.Checked = false; CheckBox_CF.Checked = false;
            CIS_CheckBox.Checked = false;

            ExcelDosyaYolu = "";
            ExcelFileName_TextBox.Hint = "Please Select Xml File";
            ExcelFileName_TextBox.Text = "";
            ExcelPage_ComboBox.Items.Clear();
            progressBar.Value = 0;
            sp_DataWord.ClearData();
            CF_DataWord.ClearData();
            EE_DataWord.ClearData();
            CIS_DataWord.ClearData();
            XML_Arrays.SP_ClearData();
            XML_Arrays.EE_ClearData();
            XML_Arrays.CF_ClearData();
            XML_Arrays.CIS_ClearData();





            if (result == DialogResult.Yes)
            {

                checkBoxS11Reel.Checked = false; checkBoxS12Reel.Checked = false; checkBoxS21Reel.Checked = false; checkBoxS22Reel.Checked = false;
                checkBoxS11Lin.Checked = false; checkBoxS12Lin.Checked = false; checkBoxS21Lin.Checked = false; checkBoxS22Lin.Checked = false;
                checkBoxS11Log.Checked = false; checkBoxS12Log.Checked = false; checkBoxS21Log.Checked = false; checkBoxS22Log.Checked = false;
                checkBoxS11SWR.Checked = false; checkBoxS22SWR.Checked = false;

                checkBoxEE.Checked = false; checkBox_EE_RI.Checked = false; checkBoxRHO.Checked = false; checkBox_EE_CF.Checked = false;
                CF_checkBox_RIRC.Checked = false; CheckBox_CF.Checked = false;
                CIS_CheckBox.Checked = false;

                ExcelDosyaYolu = "";
                ExcelPage_ComboBox.Items.Clear();
                ExcelFileName_TextBox.Hint = "Please Select Xml File";
                ExcelFileName_TextBox.Text = "";
                progressBar.Value = 0;
                sp_DataWord.ClearData();
                CF_DataWord.ClearData();
                EE_DataWord.ClearData();
                CIS_DataWord.ClearData();
                XML_Arrays.SP_ClearData();
                XML_Arrays.EE_ClearData();
                XML_Arrays.CF_ClearData();
                XML_Arrays.CIS_ClearData();




            }
            else if (result == DialogResult.No)
            {

            }

            #endregion



        }

        public void SaveBasarim()
        {
            LabelProgress.Visible = false;
            Thread.Sleep(5);
            progressBar.Value = 0;
            for (int i = 0; i < 100; i++)
            {
                progressBar.Value += 1;
            }
            LabelProgress.Visible = true;
            LabelProgress.ForeColor = System.Drawing.Color.Green;
            LabelProgress.Text = @"Save successful";
        }

        private void CreateCertificate_Button_Click(object sender, EventArgs e)
        {
            LabelProgress.Visible = false;
            progressBar.Value = 0;

            try
            {
                using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                {
                    // Diyalog kutusunun başlık metni
                    folderDialog.Description = "Please select a folder to save the Word file";

                    // Eğer kullanıcı bir klasör seçerse
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Seçilen klasörün yolu TextBox'a yazdırılır.
                        LabelProgress.Visible = true;
                        LabelProgress.ForeColor = System.Drawing.Color.Green;
                        LabelProgress.Text = @"Folder selection successful";

                        WordFolderPath = folderDialog.SelectedPath;

                    }
                }

                createTemplate.ResultPages(tables);


                if (tables.Count >= 1)
                {

                    WordBasarim();
                }
                else
                {
                    MessageBox.Show(@"Please select at least one parameter!", @"ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LabelProgress.Visible = true;
                    LabelProgress.ForeColor = System.Drawing.Color.Red;
                    LabelProgress.Text = @"ERROR!: Parameter Select";
                }

            }
            catch (Exception err)
            {
                LabelProgress.Visible = true;
                LabelProgress.ForeColor = System.Drawing.Color.Red;
                LabelProgress.Text = @"ERROR!: Word";
                MessageBox.Show(err.Message, err.StackTrace, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            try
            {
                using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                {
                    // Diyalog kutusunun başlık metni
                    folderDialog.Description = "Please select a folder to save the XML file.";

                    // Eğer kullanıcı bir klasör seçerse
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Seçilen klasörün yolu TextBox'a yazdırılır.
                        LabelProgress.Visible = true;
                        LabelProgress.ForeColor = System.Drawing.Color.Green;
                        LabelProgress.Text = @"Folder selection successful";

                        XMLFolderPath = folderDialog.SelectedPath;
                    }
                }

                string xmlSavePath = XMLFolderPath + "\\" + ExcelDosyaAdi + ".xml";
                xml.Save(xmlSavePath);
                XMLBasarim();

            }
            catch (Exception err)
            {
                LabelProgress.Visible = true;
                LabelProgress.ForeColor = System.Drawing.Color.Red;
                LabelProgress.Text = @"ERROR!: XML";
                MessageBox.Show(err.Message, err.StackTrace, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        #endregion
        public void WordBasarim()
        {
            LabelProgress.Visible = false;
            Thread.Sleep(10);
            progressBar.Value = 0;
            for (int i = 0; i < 100; i++)
            {
                progressBar.Value += 1;
            }
            LabelProgress.Visible = true;
            LabelProgress.ForeColor = System.Drawing.Color.Green;
            LabelProgress.Text = @"Transfer to word successful";
        }
        public void XMLBasarim()
        {
            LabelProgress.Visible = false;
            Thread.Sleep(10);
            progressBar.Value = 0;
            for (int i = 0; i < 100; i++)
            {
                progressBar.Value += 1;
            }
            LabelProgress.Visible = true;
            LabelProgress.ForeColor = System.Drawing.Color.Green;
            LabelProgress.Text = @"Transfer to XML successful";
        }
        private string GetExcelColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private void DisplayExcelWorksheet(string worksheetName)
        {
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(openFileDialog.FileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName];
                if (worksheet == null)
                {
                    MessageBox.Show("Selected worksheet could not be loaded.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                // Excel'deki sütun başlıklarını DataGridView'e aktar
                for (int i = 1; i <= colCount; i++)
                {
                    string columnName = GetExcelColumnName(i);
                    dataGridView1.Columns.Add(columnName, columnName);
                }

                // Excel'deki satır başlıklarını ve verileri DataGridView'e aktar
                for (int i = 1; i <= rowCount; i++)
                {
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(dataGridView1);

                    // Satır başlığı olarak satır indeksini kullan
                    row.HeaderCell.Value = i.ToString();

                    for (int j = 1; j <= colCount; j++)
                    {
                        object value = worksheet.Cells[i, j].Value;
                        row.Cells[j - 1].Value = value != null ? value.ToString() : "";
                    }

                    dataGridView1.Rows.Add(row);
                }

                // DataGridView'e sütun başlıklarını gösterme
                dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
        }

        private void CertificateForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1 || e.ColumnIndex == -1)
                return;

            int columnIndex = e.ColumnIndex;
            int rowIndex = e.RowIndex;
            object cellValue = dataGridView1.Rows[rowIndex].Cells[columnIndex].Value;
            string columnName = dataGridView1.Columns[columnIndex].HeaderText;
            int rowNumber = rowIndex + 1;

            sütun = columnName;
            satır = rowNumber;


            label4.Text = ($"Selection cell:  {"Column: "}{columnName}{"  Row: "}{rowNumber}");
            LabelProgress.Text = "Cell selection successful";





        }


        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
