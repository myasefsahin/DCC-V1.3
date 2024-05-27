

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
using System.Text.Json;
using CheckBox = System.Windows.Forms.CheckBox;
using Control = System.Windows.Forms.Control;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Bibliography;


namespace DCC
{

    public partial class CertificateForm : Form
    {

        //Constructor Metod
        public CertificateForm()
        {
            InitializeComponent();
            CheckBoxTabpagecontrol();
            RFPowtabpageControl();
            ExcelPage_ComboBox.Enabled = false;
            SelectExcel_Button.Enabled = false;


        }
        //Form kapatıldığında programın sonlanmasını sağlayan kod parçası 
        private void CertificateForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }


        //Form yüklendiğinde çalışanlar
        private void CertificateForm_Load(object sender, EventArgs e)
        {
            xml.Load(xmlFilePath);
            CheckBoxTabpagecontrol();
            RFPowtabpageControl();
            ExcelPage_ComboBox.Enabled = false;
            SelectExcel_Button.Enabled = false;
            LaboratoryComboBox.Items.Add("RF ve Mikrodalga Laboratuvarı");
            Measurement_Unc_TextBox.Text = "                                                                                                                                                                                                                                             Beyan edilen genişletilmiş ölçüm belirsizliği, standart ölçüm belirsizliğinin normal dağılım için yaklaşık  " +
            "  % 95 güvenilirlik seviyesini sağlayan k = 2 kapsam faktörü ile çarpımının sonucudur. Standart ölçüm belirsizliği GUM ve EA-4/02 dokümanlarına uygun " +
            "olarak belirlenmiştir.";
            LaboratoryComboBox.Enabled = false;
            DeviceNameTextBox.Enabled = false;
            ModelNameTextBox.Enabled = false;
            SerialNumberTextBox.Enabled = false;
            CalCodeTextBox.Enabled = false;
            SelectDeviceButton.Enabled = false;
            Instrument_DeviceTextBox.Enabled = false;
            Cal_MethodTextBox.Enabled = false;
            Measurement_Unc_TextBox.Enabled = false;
            Comments_TextBox.Enabled = false;
            ReceiveData_Button.Enabled = false;
            CreateCertificate_Button.Enabled = false;
            BackBox3.Enabled = true;



        }

        #region Tanımlamalar

        //Kullanılan değişkenler ve nesneler
        int listboxnum;
        string xmlFilePath = "xmlData/dcc_xml_scheme_uzun.xml";     
        OpenFileDialog openFileDialog = new OpenFileDialog();
        public string ExcelDosyaYolu;
        public string ExcelDosyaAdi;
        public string WordFolderPath;
        string pageName;
        public List<Table> tables = new List<Table>();
        public string XMLFolderPath;
        XmlDocument xml = new XmlDocument();
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
        CreateTable createtable = new CreateTable();
        Noise_DataWord Noise_DataWord = new Noise_DataWord();
        Noise_WordTable Noise_WordTable = new Noise_WordTable();
        Absolute_RF_Power_DataWord Absolute_RF_Power = new Absolute_RF_Power_DataWord();
        AbsoluteRF_Power_Word_Table Absolute_WordTable = new AbsoluteRF_Power_Word_Table();
        RF_Difference_DataWord RF_Difference_DataWord = new RF_Difference_DataWord();
        RF_Difference_WordTable RF_Difference_wordTable = new RF_Difference_WordTable();
        RF_Gain_DataWord RF_Gain_DataWord = new RF_Gain_DataWord();
        RF_Gain_WordTable RF_Gain_WordTable = new RF_Gain_WordTable();
        int sayac = 0;
        int satır;
        string sütun;



        #endregion

        #region API PAGE
        private async void SelectDeviceButton_Click(object sender, EventArgs e)
        {
            Instrument_DeviceTextBox.Enabled = true;
            string apiUrl = "https://localhost:7166/AdministrativeData/GetAdministrativeDataSipNo?dataId=" + OrderNumberTextBox.Text;

            using (var httpClient = new HttpClient())
            {
                try
                {
                    var response = await httpClient.GetAsync(apiUrl);
                    response.EnsureSuccessStatusCode();

                    using (var responseStream = await response.Content.ReadAsStreamAsync())
                    {
                        var options = new JsonSerializerOptions
                        {
                            PropertyNameCaseInsensitive = true
                        };

                        var responseData = await JsonSerializer.DeserializeAsync<ApiResponse>(responseStream, options);


                    }
                }
                catch (HttpRequestException ex)
                {
                    Console.WriteLine($"HTTP request error: {ex.Message}");
                }

            }


        }


        private void OrderNumberTextBox_TextChanged(object sender, EventArgs e)
        {
            LaboratoryComboBox.Enabled = !string.IsNullOrWhiteSpace(OrderNumberTextBox.Text);
            if (!LaboratoryComboBox.Enabled)
            {
                DeviceNameTextBox.Enabled = false;
                ModelNameTextBox.Enabled = false;
                SerialNumberTextBox.Enabled = false;
                CalCodeTextBox.Enabled = false;
                SelectDeviceButton.Enabled = false;
            }
        }

        private void LaboratoryComboBox_TextChanged(object sender, EventArgs e)
        {
            DeviceNameTextBox.Enabled = !string.IsNullOrWhiteSpace(LaboratoryComboBox.Text);
            if (!DeviceNameTextBox.Enabled)
            {
                ModelNameTextBox.Enabled = false;
                SerialNumberTextBox.Enabled = false;
                CalCodeTextBox.Enabled = false;
                SelectDeviceButton.Enabled = false;
            }
        }

        private void DeviceNameTextBox_TextChanged(object sender, EventArgs e)
        {
            ModelNameTextBox.Enabled = !string.IsNullOrWhiteSpace(DeviceNameTextBox.Text);
            if (!ModelNameTextBox.Enabled)
            {
                SerialNumberTextBox.Enabled = false;
                CalCodeTextBox.Enabled = false;
                SelectDeviceButton.Enabled = false;
            }
        }

        private void ModelNameTextBox_TextChanged(object sender, EventArgs e)
        {
            SerialNumberTextBox.Enabled = !string.IsNullOrWhiteSpace(ModelNameTextBox.Text);
            if (!SerialNumberTextBox.Enabled)
            {
                CalCodeTextBox.Enabled = false;
                SelectDeviceButton.Enabled = false;
            }
        }

        private void SerialNumberTextBox_TextChanged(object sender, EventArgs e)
        {
            CalCodeTextBox.Enabled = !string.IsNullOrWhiteSpace(SerialNumberTextBox.Text);
            if (!CalCodeTextBox.Enabled)
            {
                SelectDeviceButton.Enabled = false;
            }
        }

        private void CalCodeTextBox_TextChanged(object sender, EventArgs e)
        {
            SelectDeviceButton.Enabled = !string.IsNullOrWhiteSpace(CalCodeTextBox.Text);
        }

        private void DeviceTextBox_TextChanged(object sender, EventArgs e)
        {
            Cal_MethodTextBox.Enabled = !string.IsNullOrWhiteSpace(Instrument_DeviceTextBox.Text);
        }

        private void MethodTextBox_TextChanged(object sender, EventArgs e)
        {
            Measurement_Unc_TextBox.Enabled = !string.IsNullOrWhiteSpace(Cal_MethodTextBox.Text);
        }

        private void CalibrationDescTextBox_TextChanged(object sender, EventArgs e)
        {
            Comments_TextBox.Enabled = !string.IsNullOrWhiteSpace(Measurement_Unc_TextBox.Text);
        }
        #endregion

        #region Next,Back Button ve Combobox Kontrolleri 

        //Form üzerindeki Next ve Back butonlarınmın kontrolleri
        // Next ve Back butonlarına tıklandığında tıklanmaya göre değişen bazı textler ve geçiş ayarlamaları bu kısımda yapılmıştır.

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

        //Bu fonksiyon MeasurementTypes_ComboBoxtaki seçilen indexe göre açılan CheckBoxTabControl sayfalarının değişmesini sağlar.
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
                CheckBoxTabControl.SelectedTab = RFPow_Page;
            }

            if (MeasurementTypes_ComboBox.SelectedIndex == 4)
            {
                CheckBoxTabControl.SelectedTab = SParam_Page;
            }
            if (MeasurementTypes_ComboBox.SelectedIndex == 5)
            {
                CheckBoxTabControl.SelectedTab = Noise_Page;
            }
            if (MeasurementTypes_ComboBox.SelectedIndex == 6)
            {
                CheckBoxTabControl.SelectedTab = MetCH_Page;
            }
        }
        #endregion

        #region Button işlemleri


        //Bu fonksiyon ExcelPage_ComboBox üzerinde seçtiği Excel dosyası sayfası ayarlamalarını yapar.
        //Sayfa seçimini DisplayExcelWorksheet fonksiyonunu çağırarak datagridview üzerine yansıtır.
        private void ExcelPage_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            BackBox3.Enabled = false;
            Save_Row_Col_Button.Enabled = false;
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
            BackBox3.Enabled = true;

        }

        //Bu fonksiyon ExcelviewPage sayfasındaki datagrid üzerinden seçilen hücrenin satır ve sütun değerlerini değişkenlere atamnaya ve bu işlemin kontrollerini sağlamaya yarar.
        //Seçime göre kullanıcıya label4 üzerinde uyarılar verir
        private void Save_Row_Col_Button_Click(object sender, EventArgs e)
        {
            if (satır == 0 && sütun == null)
            {
                label4.Location = new System.Drawing.Point(475, 53);
                BackBox3.Enabled = false;
                label4.Text = "Please select a cell";
            }
            else
            {
                sütun = sütun.ToUpper();
                this.Text = "DATA OPERATIONS";
                CertificateTabControl.SelectedTab = DATA_PAGE;
                label4.Text = "Please double click on the cell to select it.";
                UnEnabledAllCheckBoxes(this);
                NextBox1.Visible = true;

            }
            progressBar.Value = 0;
            LabelProgress.Text = "";
            ExcelPage_ComboBox.Enabled = false;

        }

        private void SelectExcel_Button_Click(object sender, EventArgs e)
        {
            {
                // Progress label'ını görünmez yap ve progress bar'ı sıfırla
                LabelProgress.Visible = false;
                progressBar.Value = 0;
                try
                {
                    // OpenFileDialog filtresini belirle ve başlangıç dizinini ayarla
                    openFileDialog.Filter = @"Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                    openFileDialog.InitialDirectory = "C:\\";
                    openFileDialog.Title = @"Excel Files, Select a "".xlsx"" file";

                    // Kullanıcı bir dosya seçtiyse
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Seçilen dosyanın yolunu ve adını al
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
                            ExcelPage_ComboBox.Enabled = true; // ComboBox'ı etkinleştir
                        }

                        // İşlem başarılı oldu mesajını göster
                        LabelProgress.Visible = true;
                        LabelProgress.ForeColor = System.Drawing.Color.Green;
                        LabelProgress.Text = @"Excel file selection successful";
                    }
                }
                catch (Exception err)
                {
                    // Hata durumunda hata mesajını göster
                    LabelProgress.Visible = true;
                    LabelProgress.ForeColor = System.Drawing.Color.Red;
                    LabelProgress.Text = @"ERROR!: Selection of Excel";
                    MessageBox.Show(err.Message, err.StackTrace, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        //Bu fonksiyon verilerin çekildiği ve formatlandığı ana fonksiyondur.     
        private void ReceiveData_Button_Click(object sender, EventArgs e)
        {
            this.xml = CreateXML.AddAdministrativeData(xml, XML_Arrays);
            try
            {

                #region S-Parametre
                if (MeasurementTypes_ComboBox.SelectedIndex == 4)//Eğer s-parametre seçili ise bu if içerisine girer ve işlemleri gerçekleştirir.
                {

                    sp_DataWord.main(ExcelDosyaYolu, pageName, satır, sütun);//Veriler bu satırda S-parametre Dataword sınıfı nesnesi oluşturularak çekilir gerekli parametreler girilir ve formatlanır.
                    XML_Arrays.SP_Data_Xml(ExcelDosyaYolu, pageName, satır, sütun);//Veriler bu satırda XML çıktısı için formatlanmadan çekilir.
                    label7.Visible = false;
                    listboxnum = listboxnum + 1;
                    listBox1.Items.Add((listboxnum) + "-) " + "S-Parametre -" + ExcelDosyaAdi + "_" + pageName);//Transfer edilen veri bu kısımda dosya adı ve hangi ölçüm adı olarak listbox içerisine eklenir.

                    #region S parametre Checkbox Kontrolleri

                    if (ExcelFileName_TextBox.Text != "(Please enter a header name..)")
                    {
                        TableName = MeasurementTypes_ComboBox.Text + " - ";
                    }
                    else
                    {
                        TableName = "";
                    }

                    // Bu liste seçilen checkboxlara göre xml oluşturmak için kullanılır.
                    //Aşağıdaki kontrolde liste elemanlarının değerleri değiştirilir ve değiştirilen liste CreateXml içerisindeki ilgili  methodlara parametre olarak gönderilir.
                    List<bool> dataList = new List<bool>(14) { false, false, false, false, false, false, false, false, false, false, false, false, false, false };

                    //Bu aşamada oluşturulan kullanıcın seçtiği checkboxlar kontrol edilir ve true veya false olmasına göre gerekli tablolar oluşturulur ve bir tablo listesine aktarılır.
                    //Seçime göre yukarıda bulunan ilgili listenin indexinin elemanı değiştirilir ve CreateXML sınıfı içeriisnde bulunan ilgili fonksiyona liste parametre olarak verilir. 

                    if (checkBoxS11Reel.Checked)
                    {
                        sayac++;//Tablo numarası sayacı
                        Table s11reelTable = SP_WordTable.CreateReelImg(ExcelDosyaAdi, sayac, sp_DataWord.tableName1, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS11Reel, sp_DataWord.ArrayS11ReelUnc, sp_DataWord.ArrayS11Complex, sp_DataWord.ArrayS11ComplexUnc);
                        tables.Add(s11reelTable);
                        dataList[0] = true; //seçime göre liste elemanı değiştirilir. 
                    }

                    if (checkBoxS11Lin.Checked)
                    {
                        sayac++;
                        Table s11linTable = SP_WordTable.CreateLinPhase(ExcelDosyaAdi, sayac, sp_DataWord.tableName2, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS11Lin, sp_DataWord.ArrayS11LinUnc, sp_DataWord.ArrayS11LinPhase, sp_DataWord.ArrayS11LinPhaseUnc);
                        tables.Add(s11linTable);
                        dataList[1] = true;
                    }
                    if (checkBoxS11Log.Checked)
                    {
                        sayac++;
                        Table s11logTable = SP_WordTable.CreateLogPhase(ExcelDosyaAdi, sayac, sp_DataWord.tableName3, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS11Log, sp_DataWord.ArrayS11LogUnc, sp_DataWord.ArrayS11LogPhase, sp_DataWord.ArrayS11LogPhaseUnc);
                        tables.Add(s11logTable);
                        dataList[2] = true;
                    }
                    if (checkBoxS11SWR.Checked)
                    {
                        sayac++;
                        Table s11swrTable = SP_WordTable.CreateSWR(ExcelDosyaAdi, sayac, sp_DataWord.tableName4, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS11SWR, sp_DataWord.ArrayS11SWRUnc);
                        tables.Add(s11swrTable);
                        dataList[3] = true;
                    }


                    if (checkBoxS12Reel.Checked)
                    {
                        sayac++;
                        Table s12reelTable = SP_WordTable.CreateReelImg(ExcelDosyaAdi, sayac, sp_DataWord.tableName5, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS12Reel, sp_DataWord.ArrayS12ReelUnc, sp_DataWord.ArrayS12Complex, sp_DataWord.ArrayS12ComplexUnc);
                        tables.Add(s12reelTable);
                        dataList[4] = true;
                    }
                    if (checkBoxS12Lin.Checked)
                    {
                        sayac++;
                        Table s12linTable = SP_WordTable.CreateLinPhase(ExcelDosyaAdi, sayac, sp_DataWord.tableName6, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS12Lin, sp_DataWord.ArrayS12LinUnc, sp_DataWord.ArrayS12LinPhase, sp_DataWord.ArrayS12LinPhaseUnc);
                        tables.Add(s12linTable);
                        dataList[5] = true;

                    }
                    if (checkBoxS12Log.Checked)
                    {
                        sayac++;
                        Table s12logTable = SP_WordTable.CreateLogPhase(ExcelDosyaAdi, sayac, sp_DataWord.tableName7, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS12Log, sp_DataWord.ArrayS12LogUnc, sp_DataWord.ArrayS12LogPhase, sp_DataWord.ArrayS12LogPhaseUnc);
                        tables.Add(s12logTable);
                        dataList[6] = true;

                    }


                    if (checkBoxS21Reel.Checked)
                    {
                        sayac++;
                        Table s21reelTable = SP_WordTable.CreateReelImg(ExcelDosyaAdi, sayac, sp_DataWord.tableName8, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS21Reel, sp_DataWord.ArrayS21ReelUnc, sp_DataWord.ArrayS21Complex, sp_DataWord.ArrayS21ComplexUnc);
                        tables.Add(s21reelTable);
                        dataList[7] = true;

                    }
                    if (checkBoxS21Lin.Checked)
                    {
                        sayac++;
                        Table s21linTable = SP_WordTable.CreateLinPhase(ExcelDosyaAdi, sayac, sp_DataWord.tableName9, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS21Lin, sp_DataWord.ArrayS21LinUnc, sp_DataWord.ArrayS21LinPhase, sp_DataWord.ArrayS21LinPhaseUnc);
                        tables.Add(s21linTable);
                        dataList[8] = true;

                    }
                    if (checkBoxS21Log.Checked)
                    {
                        sayac++;
                        Table s21logTable = SP_WordTable.CreateLogPhase(ExcelDosyaAdi, sayac, sp_DataWord.tableName10, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS21Log, sp_DataWord.ArrayS21LogUnc, sp_DataWord.ArrayS21LogPhase, sp_DataWord.ArrayS21LogPhaseUnc);
                        tables.Add(s21logTable);
                        dataList[9] = true;

                    }

                    if (checkBoxS22Reel.Checked)
                    {
                        sayac++;
                        Table s22reelTable = SP_WordTable.CreateReelImg(ExcelDosyaAdi, sayac, sp_DataWord.tableName11, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS22Reel, sp_DataWord.ArrayS22ReelUnc, sp_DataWord.ArrayS22Complex, sp_DataWord.ArrayS22ComplexUnc);
                        tables.Add(s22reelTable);
                        dataList[10] = true;

                    }
                    if (checkBoxS22Lin.Checked)
                    {
                        sayac++;
                        Table s22linTable = SP_WordTable.CreateLinPhase(ExcelDosyaAdi, sayac, sp_DataWord.tableName12, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS22Lin, sp_DataWord.ArrayS22LinUnc, sp_DataWord.ArrayS22LinPhase, sp_DataWord.ArrayS22LinPhaseUnc);
                        tables.Add(s22linTable);
                        dataList[11] = true;

                    }
                    if (checkBoxS22Log.Checked)
                    {
                        sayac++;
                        Table s22logTable = SP_WordTable.CreateLogPhase(ExcelDosyaAdi, sayac, sp_DataWord.tableName13, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS22Log, sp_DataWord.ArrayS22LogUnc, sp_DataWord.ArrayS22LogPhase, sp_DataWord.ArrayS22LogPhaseUnc);
                        tables.Add(s22logTable);
                        dataList[12] = true;

                    }
                    if (checkBoxS22SWR.Checked)
                    {
                        sayac++;
                        Table s22swrTable = SP_WordTable.CreateSWR(ExcelDosyaAdi, sayac, sp_DataWord.tableName14, sp_DataWord.ArrayFrekans, sp_DataWord.ArrayS22SWR, sp_DataWord.ArrayS22SWRUnc);
                        tables.Add(s22swrTable);
                        dataList[13] = true;

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
                    label7.Visible = false;
                    listboxnum = listboxnum + 1;
                    listBox1.Items.Add((listboxnum) + "-) " + "Effective Efficiency -" + ExcelDosyaAdi + "_" + pageName);

                    #region EE ChechkBox Kontrolleri
                    List<bool> dataListEE = new List<bool>(14) { false, false, false, false };


                    if (checkBoxEE.Checked)
                    {
                        sayac++;
                        Table s11reelTable_EE = EE_WordTable.EECreateReelImg(ExcelDosyaAdi, sayac, EE_DataWord.tableName1, EE_DataWord.EE_ArrayFrekans, EE_DataWord.EE_ArrayS11Reel, EE_DataWord.EE_ArrayS11ReelUnc, EE_DataWord.EE_ArrayS11Complex, EE_DataWord.EE_ArrayS11ComplexUnc);
                        tables.Add(s11reelTable_EE);
                        dataListEE[0] = true;

                    }

                    if (checkBox_EE_RI.Checked)
                    {
                        sayac++;
                        Table EE_table = EE_WordTable.CreateEE(ExcelDosyaAdi, sayac, EE_DataWord.tableName2, EE_DataWord.EE_ArrayFrekans, EE_DataWord.EE_ArrayEE, EE_DataWord.EE_ArrayEEUnc);
                        tables.Add(EE_table);
                        dataListEE[1] = true;

                    }

                    // Logarithmic Kutusu
                    if (checkBoxRHO.Checked)
                    {
                        sayac++;
                        Table RHOTable = EE_WordTable.CreateRHO(ExcelDosyaAdi, sayac, EE_DataWord.tableName3, EE_DataWord.EE_ArrayFrekans, EE_DataWord.EE_ArrayRhoLin, EE_DataWord.EE_ArrayRhoUnc);
                        tables.Add(RHOTable);
                        dataListEE[2] = true;

                    }
                    if (checkBox_EE_CF.Checked)
                    {
                        sayac++;
                        Table EE_CF_table = EE_WordTable.CreateCF(ExcelDosyaAdi, sayac, EE_DataWord.tableName4, EE_DataWord.EE_ArrayFrekans, EE_DataWord.EE_ArrayCF, EE_DataWord.EE_ArrayCFUnc);
                        tables.Add(EE_CF_table);
                        dataListEE[3] = true;

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
                    label7.Visible = false;
                    listBox1.Items.Add((listboxnum + 1) + "-) " + "Calibration Factor -" + ExcelDosyaAdi + "_" + pageName);

                    #region CF CheckBox Kontrolleri

                    List<bool> dataListCF = new List<bool>(14) { false, false };

                    if (CheckBox_CF.Checked)
                    {
                        sayac++;
                        Table reelTable_CF = CF_Word_Table.CF_CreateReelImg(ExcelDosyaAdi, sayac, CF_DataWord.tableName1, CF_DataWord.CF_ArrayFrekans, CF_DataWord.CF_ArrayReel, CF_DataWord.CF_ArrayReelUnc, CF_DataWord.CF_ArrayComplex, CF_DataWord.CF_ArrayComplexUnc, CF_DataWord.CF_YK, CF_DataWord.CF_YK_Unc);
                        tables.Add(reelTable_CF);
                        dataListCF[0] = true;

                    }


                    if (CF_checkBox_RIRC.Checked)
                    {
                        sayac++;
                        Table CF_table = CF_Word_Table.CF_CreateCF(ExcelDosyaAdi, sayac, CF_DataWord.tableName2, CF_DataWord.CF_ArrayFrekans, CF_DataWord.CF_Array, CF_DataWord.CF_ArrayCFUnc);
                        tables.Add(CF_table);
                        dataListCF[1] = true;

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
                    label7.Visible = false;
                    listboxnum = listboxnum + 1;
                    listBox1.Items.Add((listboxnum) + "-) " + "Calculable Impedance Standard -" + ExcelDosyaAdi + "_" + pageName);

                    bool CIS_bool = false;


                    if (CIS_CheckBox.Checked)
                    {
                        sayac++;

                        Table CIS_table = CIS_Word_Table.Create_Z_Position(ExcelDosyaAdi, sayac, CIS_DataWord.tableName, CIS_DataWord.CIS_Olcum_Adım, CIS_DataWord.CIS_ZP, CIS_DataWord.CIS_ZP_Unc, CIS_DataWord.CIS_ICOD, CIS_DataWord.CIS_ICOD_Unc, CIS_DataWord.CIS_OCID, CIS_DataWord.CIS_OCID_Unc);
                        tables.Add(CIS_table);

                        CIS_bool = true;

                    }
                    CreateXML.AddCISResult(xml, TableName, XML_Arrays, CIS_bool);

                }
                #endregion

                #region Noise
                else if (MeasurementTypes_ComboBox.SelectedIndex == 5)
                {

                    Noise_DataWord.main(ExcelDosyaYolu, pageName, satır, sütun);
                    XML_Arrays.Noise_Data_Xml(ExcelDosyaYolu, pageName, satır, sütun);
                    label7.Visible = false;
                    listboxnum = listboxnum + 1;
                    listBox1.Items.Add((listboxnum) + "-) " + "Noise -" + ExcelDosyaAdi + "_" + pageName);


                    List<bool> NoiseBool = new List<bool>(3) { false, false, false };

                    if (NS_checkBoxENR.Checked)
                    {
                        sayac++;
                        Table ENR_Noise_table = Noise_WordTable.CreateENR(ExcelDosyaAdi, sayac, Noise_DataWord.tableName1, Noise_DataWord.NS_ArrayFrekans, Noise_DataWord.NS_ArrayENR, Noise_DataWord.NS_ArrayENRUnc);
                        tables.Add(ENR_Noise_table);
                        NoiseBool[0] = true;

                    }
                    if (NS_checkBox_DC_ON.Checked)
                    {
                        sayac++;
                        Table DC_ON_Noise_table = Noise_WordTable.Create_DC_ON_OFF(ExcelDosyaAdi, sayac, Noise_DataWord.tableName2, Noise_DataWord.NS_ArrayFrekans, Noise_DataWord.NS_ArrayRC, Noise_DataWord.NS_ArrayRC_ustlimit, Noise_DataWord.NS_ArrayRCUnc,
                                                                                    Noise_DataWord.NS_ArrayRC_Phase, Noise_DataWord.NS_ArrayRC_PhaseUnc);
                        tables.Add(DC_ON_Noise_table);
                        NoiseBool[1] = true;

                    }

                    if (NS_checkBox_DC_OFF.Checked)
                    {
                        sayac++;
                        Table DC_OFF_Noise_table = Noise_WordTable.Create_DC_ON_OFF(ExcelDosyaAdi, sayac, Noise_DataWord.tableName3, Noise_DataWord.NS_ArrayFrekans, Noise_DataWord.NS_ArrayRC_DC_OFF, Noise_DataWord.NS_ArrayRC_ustlimit_DC_OFF, Noise_DataWord.NS_ArrayRCUnc_DC_OFF,
                                                                                    Noise_DataWord.NS_ArrayRC_Phase_DC_OFF, Noise_DataWord.NS_ArrayRC_PhaseUnc_DC_OFF);
                        tables.Add(DC_OFF_Noise_table);
                        NoiseBool[2] = true;

                    }
                    CreateXML.AddNoiseResult(xml, TableName, XML_Arrays, NoiseBool);

                }
                #endregion

                #region Absolute RF Power
                else if (RFPowTabControl.SelectedTab == Abs_RF_Power_tabpage)
                {

                    Absolute_RF_Power.main(ExcelDosyaYolu, pageName, satır, sütun);
                    XML_Arrays.ABS_RFP_Data_Xml(ExcelDosyaYolu, pageName, satır, sütun);
                    label7.Visible = false;
                    listboxnum = listboxnum + 1;
                    listBox1.Items.Add((listboxnum) + "-) " + "Absolute RF Power -" + ExcelDosyaAdi + "_" + pageName);


                    List<bool> ARFPBool = new List<bool>(3) { false, false, false, false, false, false, false, false, false, false, false };

                    if (ARFP_1.Checked)
                    {
                        sayac++;
                        Table ARFP1_table = Absolute_WordTable.ARFP_CreateTable_1(ExcelDosyaAdi, sayac, Absolute_RF_Power.tableName1, Absolute_RF_Power.ARFP_T1_Frekans, Absolute_RF_Power.ARFP_T1_Cıkıs_Gücü, Absolute_RF_Power.ARFP_T1_Olculen_Güc, Absolute_RF_Power.ARFP_T1_AltSınır,
                                                                                  Absolute_RF_Power.ARFP_T1_Sapma, Absolute_RF_Power.ARFP_T1_ÜstSınır, Absolute_RF_Power.ARFP_T1_Belirsizlik, "Ölçülen Güç (dBm)", "Sapma");
                        tables.Add(ARFP1_table);
                        ARFPBool[0] = true;

                    }

                    if (ARFP_2.Checked)
                    {
                        sayac++;
                        Table ARFP2_table = Absolute_WordTable.ARFP_CreateTable_1(ExcelDosyaAdi, sayac, Absolute_RF_Power.tableName2, Absolute_RF_Power.ARFP_T2_Frekans, Absolute_RF_Power.ARFP_T2_Cıkıs_Gücü, Absolute_RF_Power.ARFP_T2_OlculenDeger, Absolute_RF_Power.ARFP_T2_AltSınır,
                                                                                  Absolute_RF_Power.ARFP_T2_Fark, Absolute_RF_Power.ARFP_T2_ÜstSınır, Absolute_RF_Power.ARFP_T2_Belirsizlik, "Ölçülen Değer (dBm)", "Fark (dB)");
                        tables.Add(ARFP2_table);
                        ARFPBool[1] = true;

                    }
                    if (ARFP_3.Checked)
                    {
                        sayac++;
                        Table ARFP3_table = Absolute_WordTable.ARFP_CreateTable_1(ExcelDosyaAdi, sayac, Absolute_RF_Power.tableName3, Absolute_RF_Power.ARFP_T3_Frekans, Absolute_RF_Power.ARFP_T3_Cıkıs_Gücü, Absolute_RF_Power.ARFP_T3_OlculenZayıflatma, Absolute_RF_Power.ARFP_T3_AltSınır,
                                                                                  Absolute_RF_Power.ARFP_T3_Zayıflatma, Absolute_RF_Power.ARFP_T3_ÜstSınır, Absolute_RF_Power.ARFP_T3_Belirsizlik, "Ölçülen Zayıflatma (dB)", "Zayıflatma Hatası (dB)");
                        tables.Add(ARFP3_table);
                        ARFPBool[2] = true;

                    }
                    if (ARFP_4.Checked)
                    {
                        sayac++;
                        Table ARFP4_table = Absolute_WordTable.ARFP_CreateTable_2(ExcelDosyaAdi, sayac, Absolute_RF_Power.tableName4, Absolute_RF_Power.ARFP_T4_T5_T6_frekans, Absolute_RF_Power.ARFP_T4_SWR_Seviye, Absolute_RF_Power.ARFP_T4_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T4_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T4_SWR_Belirsizlik, "Maksimum Değer");
                        tables.Add(ARFP4_table);
                        ARFPBool[3] = true;

                    }
                    if (ARFP_5.Checked)
                    {
                        sayac++;
                        Table ARFP5_table = Absolute_WordTable.ARFP_CreateTable_2(ExcelDosyaAdi, sayac, Absolute_RF_Power.tableName5, Absolute_RF_Power.ARFP_T4_T5_T6_frekans, Absolute_RF_Power.ARFP_T5_SWR_Seviye, Absolute_RF_Power.ARFP_T5_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T5_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T5_SWR_Belirsizlik, "Maksimum Değer");
                        tables.Add(ARFP5_table);
                        ARFPBool[4] = true;

                    }
                    if (ARFP_6.Checked)
                    {
                        sayac++;
                        Table ARFP6_table = Absolute_WordTable.ARFP_CreateTable_2(ExcelDosyaAdi, sayac, Absolute_RF_Power.tableName6, Absolute_RF_Power.ARFP_T4_T5_T6_frekans, Absolute_RF_Power.ARFP_T6_SWR_Seviye, Absolute_RF_Power.ARFP_T6_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T6_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T6_SWR_Belirsizlik, "Maksimum Değer");
                        tables.Add(ARFP6_table);
                        ARFPBool[5] = true;

                    }
                    if (ARFP_7.Checked)
                    {
                        sayac++;
                        Table ARFP7_table = Absolute_WordTable.ARFP_CreateTable_1(ExcelDosyaAdi, sayac, Absolute_RF_Power.tableName7, Absolute_RF_Power.ARFP_T7_Frekans, Absolute_RF_Power.ARFP_T7_Cıkıs_Gücü, Absolute_RF_Power.ARFP_T7_OlculenGuc, Absolute_RF_Power.ARFP_T7_AltSınır,
                                                                                  Absolute_RF_Power.ARFP_T7_Sapma, Absolute_RF_Power.ARFP_T7_ÜstSınır, Absolute_RF_Power.ARFP_T7_Belirsizlik, "Ölçülen Güç (dBm)", "Sapma(dB)");
                        tables.Add(ARFP7_table);
                        ARFPBool[6] = true;

                    }
                    if (ARFP_8.Checked)
                    {
                        sayac++;
                        Table ARFP8_table = Absolute_WordTable.ARFP_CreateTable_1(ExcelDosyaAdi, sayac, Absolute_RF_Power.tableName8, Absolute_RF_Power.ARFP_T8_Frekans, Absolute_RF_Power.ARFP_T8_Cıkıs_Gücü, Absolute_RF_Power.ARFP_T8_OlculenDeger, Absolute_RF_Power.ARFP_T8_AltSınır,
                                                                                  Absolute_RF_Power.ARFP_T8_Fark, Absolute_RF_Power.ARFP_T8_ÜstSınır, Absolute_RF_Power.ARFP_T8_Belirsizlik, "Ölçülen Değer (dBm)", "Fark (dB)");
                        tables.Add(ARFP8_table);
                        ARFPBool[7] = true;

                    }
                    if (ARFP_9.Checked)
                    {
                        sayac++;
                        Table ARFP9_table = Absolute_WordTable.ARFP_CreateTable_2(ExcelDosyaAdi, sayac, Absolute_RF_Power.tableName9, Absolute_RF_Power.ARFP_T9_T10_T11_frekans, Absolute_RF_Power.ARFP_T9_SWR_Seviye, Absolute_RF_Power.ARFP_T9_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T9_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T9_SWR_Belirsizlik, "Üst Sınır");
                        tables.Add(ARFP9_table);
                        ARFPBool[8] = true;

                    }
                    if (ARFP_10.Checked)
                    {
                        sayac++;
                        Table ARFP10_table = Absolute_WordTable.ARFP_CreateTable_2(ExcelDosyaAdi, sayac, Absolute_RF_Power.tableName10, Absolute_RF_Power.ARFP_T9_T10_T11_frekans, Absolute_RF_Power.ARFP_T10_SWR_Seviye, Absolute_RF_Power.ARFP_T10_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T10_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T10_SWR_Belirsizlik, "Üst Sınır");
                        tables.Add(ARFP10_table);
                        ARFPBool[9] = true;

                    }
                    if (ARFP_11.Checked)
                    {
                        sayac++;
                        Table ARFP11_table = Absolute_WordTable.ARFP_CreateTable_2(ExcelDosyaAdi, sayac, Absolute_RF_Power.tableName11, Absolute_RF_Power.ARFP_T9_T10_T11_frekans, Absolute_RF_Power.ARFP_T11_SWR_Seviye, Absolute_RF_Power.ARFP_T11_SWR_OlculenDeger, Absolute_RF_Power.ARFP_T11_SWR_MaksimumDeger, Absolute_RF_Power.ARFP_T11_SWR_Belirsizlik, "Üst Sınır");
                        tables.Add(ARFP11_table);
                        ARFPBool[10] = true;

                    }
                    CreateXML.Add_ARFP_Result(xml, TableName, XML_Arrays, ARFPBool);

                }



                #endregion

                #region RF Difference
                else if (RFPowTabControl.SelectedTab == RF_Diff_tabpage)
                {

                    RF_Difference_DataWord.main(ExcelDosyaYolu, pageName, satır, sütun);
                    XML_Arrays.RF_Diff_DataXml(ExcelDosyaYolu, pageName, satır, sütun);
                    label7.Visible = false;
                    listboxnum = listboxnum + 1;
                    listBox1.Items.Add((listboxnum) + "-) " + "RF Difference -" + ExcelDosyaAdi + "_" + pageName);


                    List<bool> RFDBool = new List<bool>(3) { false, false, false, false };

                    if (RF_Diff_1.Checked = true)
                    {

                        sayac++;
                        Table RFD1_table = RF_Difference_wordTable.RF_Diff_Table(ExcelDosyaAdi, sayac, RF_Difference_DataWord.tableName1, RF_Difference_DataWord.RFD_T1_Frekans, RF_Difference_DataWord.RFD_T1_GostergeDegeri, RF_Difference_DataWord.RFD_T1_AltSınır, RF_Difference_DataWord.RFD_T1_OlculenDeger, RF_Difference_DataWord.RFD_T1_OlculenFark,
                                             RF_Difference_DataWord.RFD_T1_ÜstSınır, RF_Difference_DataWord.RFD_T1_Belirsizlik, "Frekans (GHz)", "Gösterge Değeri (dBm)", "Alt Sınır (dBm)", "Ölçülen Değer (dBm)", "Ölçülen Fark (dB)", "Üst Sınır (dBm)", "Belirsizlik (dB)");
                        tables.Add(RFD1_table);
                        RFDBool[0] = true;

                    }
                    if (RF_Diff_2.Checked = true)
                    {

                        sayac++;
                        Table RFD2_table = RF_Difference_wordTable.RF_Diff_Table(ExcelDosyaAdi, sayac, RF_Difference_DataWord.tableName2, RF_Difference_DataWord.RFD_T2_Frekans, RF_Difference_DataWord.RFD_T2_Nom_Guc_Lvl, RF_Difference_DataWord.RFD_T2_OlculenDeger, RF_Difference_DataWord.RFD_T2_AltSınır, RF_Difference_DataWord.RFD_T2_Nom_Guc_Lvl_fark,
                                             RF_Difference_DataWord.RFD_T2_ÜstSınır, RF_Difference_DataWord.RFD_T2_Belirsizlik, "Frekans (GHz))", "Nominal Güç Seviyesi(dBm)", "Ölçülen Değer (dBm)", "Ölçülen Değer (dBm)", "Nominal Güç Seviye Farkı (dB)", "Üst Sınır (dB)", "Belirsizlik (dB)");
                        tables.Add(RFD2_table);
                        RFDBool[1] = true;

                    }

                    if (RF_Diff_3.Checked = true)
                    {
                        sayac++;
                        Table RFD3_table = RF_Difference_wordTable.RF_Diff_Table(ExcelDosyaAdi, sayac, RF_Difference_DataWord.tableName3, RF_Difference_DataWord.RFD_T3_Frekans, RF_Difference_DataWord.RFD_T3_NominalGuc, RF_Difference_DataWord.RFD_T3_AltSınır, RF_Difference_DataWord.RFD_T3_OlculenDeger, RF_Difference_DataWord.RFD_T3_ÜstSınır,
                                             RF_Difference_DataWord.RFD_T3_Fark, RF_Difference_DataWord.RFD_T3_Belirsizlik, "Frekans", "Nominal Güç (dBm)", "Alt sınır (dBm)", "Ölçülen Değer (dBm)", "Üst Sınır (dBm)", "Fark(dB)", "Belirsizlik (dB)");
                        tables.Add(RFD3_table);
                        RFDBool[2] = true;

                    }
                    if (RF_Diff_4.Checked = true)
                    {
                        sayac++;
                        Table RFD3_table = RF_Difference_wordTable.RF_Diff_Table(ExcelDosyaAdi, sayac, RF_Difference_DataWord.tableName4, RF_Difference_DataWord.RFD_T4_Min_Guc_lvl, RF_Difference_DataWord.RFD_T4_Max_Guc_lvl, RF_Difference_DataWord.RFD_T4_Frekans, RF_Difference_DataWord.RFD_T4_AltSınır, RF_Difference_DataWord.RFD_T4_Fark,
                                             RF_Difference_DataWord.RFD_T4_UstSınır, RF_Difference_DataWord.RFD_T4_Belirsizlik, "Min.Güç Seviyesi (dBm)", "Max. Güç Seviyesi (dBm)", "Frekans", "Alt Sınır (dB)", "Fark(dB)", "Üst Sınır (dB)", "Belirsizlik (dB)");
                        tables.Add(RFD3_table);
                        RFDBool[3] = true;

                    }
                    CreateXML.Add_RFD_result(xml, TableName, XML_Arrays, RFDBool);

                }
                #endregion

                #region RF GAİN

                else if (RFPowTabControl.SelectedTab == RF_Gain_tabpage)
                {
                    RF_Gain_DataWord.main(ExcelDosyaYolu, pageName, satır, sütun);
                    XML_Arrays.RF_Gain_DataXml(ExcelDosyaYolu, pageName, satır, sütun);
                    label7.Visible = false;
                    listboxnum = listboxnum + 1;
                    listBox1.Items.Add((listboxnum ) + "-) " + "RF Gain -" + ExcelDosyaAdi + "_" + pageName);



                    List<bool> RFGBool = new List<bool>(3) { false, false, false, false };

                    if (RF_Gain1.Checked = true)
                    {
                        sayac++;
                        Table RFG1_table = RF_Gain_WordTable.RF_Gain_Table(ExcelDosyaAdi, sayac, RF_Gain_DataWord.tableName1, RF_Gain_DataWord.RFG_T1_Frekans, RF_Gain_DataWord.RFG_T1_GirisGucu, RF_Gain_DataWord.RFG_T1_Belirsizlik, "Frekans", "Giriş Gücü", "Alt Sınır Belirsizlik (dB)");
                        tables.Add(RFG1_table);
                        RFGBool[0] = true;

                    }
                    if (RF_Gain2.Checked = true)
                    {
                        sayac++;
                        Table RFG2_table = RF_Gain_WordTable.RF_Gain_Table(ExcelDosyaAdi, sayac, RF_Gain_DataWord.tableName2, RF_Gain_DataWord.RFG_T2_EnBuyukKazanc, RF_Gain_DataWord.RFG_T2_EnKucukKazanc, RF_Gain_DataWord.RFG_T2_Flatness, "En Büyük Kazanç (dB)", "En Küçük Kazanç (dB)", "Flatness (±dB)");
                        tables.Add(RFG2_table);
                        RFGBool[1] = true;

                    }
                    if (RF_Gain3.Checked = true)
                    {
                        sayac++;
                        Table RFG3_table = RF_Gain_WordTable.RF_Gain_Table(ExcelDosyaAdi, sayac, RF_Gain_DataWord.tableName3, RF_Gain_DataWord.RFG_T3_Nom_Giris_Gucu, RF_Gain_DataWord.RFG_T3_Kazanc, RF_Gain_DataWord.RFG_T3_Belirsizlik, "Nominal Giriş Gücü", "Kazanç", "Uncertainty");
                        tables.Add(RFG3_table);
                        RFGBool[2] = true;

                    }
                    if (RF_Gain4.Checked = true)
                    {
                        sayac++;
                        Table RFG3_table = RF_Gain_WordTable.RF_Gain_Table(ExcelDosyaAdi, sayac, RF_Gain_DataWord.tableName4, RF_Gain_DataWord.RFG_T4_Nom_Giris_Gucu, RF_Gain_DataWord.RFG_T4_Kazanc, RF_Gain_DataWord.RFG_T4_Belirsizlik, "Nominal Giriş Gücü", "Kazanç", "Uncertainty");
                        tables.Add(RFG3_table);
                        RFGBool[3] = true;

                    }
                    CreateXML.Add_RFG_result(xml, TableName, XML_Arrays, RFGBool);


                }
                #endregion

                ImportDataBasarim();

            }

            catch (Exception err)
            {
                LabelProgress.Visible = true;
                LabelProgress.ForeColor = System.Drawing.Color.Red;
                LabelProgress.Text = @"ERROR!: Excel";
                MessageBox.Show(err.Message, err.StackTrace, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            NextBox1.Visible = false;
            refresh();



        }

        public void refresh2()
        {
            UncheckAllCheckBoxes(this);
            ExcelDosyaYolu = "";
            ExcelPage_ComboBox.Items.Clear();
            ExcelPage_ComboBox.Refresh();
            ExcelFileName_TextBox.Hint = "Please Select Xml File";
            ExcelFileName_TextBox.Text = "";
            progressBar.Value = 0;
            sp_DataWord.ClearData();
            CF_DataWord.ClearData();
            EE_DataWord.ClearData();
            CIS_DataWord.ClearData();
            Noise_DataWord.ClearData();
            Absolute_RF_Power.ClearData();
            RF_Difference_DataWord.ClearData();
            RF_Gain_DataWord.ClearData();
            XML_Arrays.SP_ClearData();
            XML_Arrays.EE_ClearData();
            XML_Arrays.CF_ClearData();
            XML_Arrays.CIS_ClearData();
            XML_Arrays.Absolute_RF_Power_ClearData();
            XML_Arrays.RF_Difference_ClearData();
            XML_Arrays.XML_RFG_ClearData();
            CreateCertificate_Button.Enabled = false;
            ReceiveData_Button.Enabled = false;
            ExcelPage_ComboBox.Enabled = false;
            EnabledAllCheckBoxes(this);
        }
        public void refresh()
        {

            DialogResult result = MessageBox.Show("Have been saved for " + MeasurementTypes_ComboBox.SelectedItem as String + " \nIf you want to add more results click Yes.\nIf not click No.", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
            {
                // Tüm CheckBox'ların işaretlerini kaldır
                UncheckAllCheckBoxes(this);

                // Excel dosya yolu ve adını sıfırla
                ExcelDosyaYolu = "";
                ExcelPage_ComboBox.Items.Clear(); // Excel sayfa isimlerini içeren ComboBox'ı temizle
                ExcelPage_ComboBox.Refresh(); // ComboBox'ı yenile
                ExcelFileName_TextBox.Hint = "Please Select Xml File"; // TextBox'a ipucu ekle
                ExcelFileName_TextBox.Text = ""; // TextBox'ı temizle

                // İlerleme çubuğunu sıfırla
                progressBar.Value = 0;

                // Veri depolarını temizle
                sp_DataWord.ClearData();
                CF_DataWord.ClearData();
                EE_DataWord.ClearData();
                CIS_DataWord.ClearData();
                Noise_DataWord.ClearData();
                Absolute_RF_Power.ClearData();
                RF_Difference_DataWord.ClearData();
                RF_Gain_DataWord.ClearData();

                // XML dizilerini temizle
                XML_Arrays.SP_ClearData();
                XML_Arrays.EE_ClearData();
                XML_Arrays.CF_ClearData();
                XML_Arrays.CIS_ClearData();
                XML_Arrays.Absolute_RF_Power_ClearData();
                XML_Arrays.RF_Difference_ClearData();
                XML_Arrays.XML_RFG_ClearData();

                // Butonları devre dışı bırak
                CreateCertificate_Button.Enabled = false;
                ReceiveData_Button.Enabled = false;
                ExcelPage_ComboBox.Enabled = false;

                // Tüm CheckBox'ları etkinleştir
                EnabledAllCheckBoxes(this);
            }

            else if (result == DialogResult.No)
            {
                CreateCertificate_Button.Enabled = true;
                sayac = 0;
                UnEnabledAllCheckBoxes(this);
                ExcelDosyaYolu = "";
                ExcelPage_ComboBox.Refresh();
                ExcelPage_ComboBox.Enabled = false;
                progressBar.Value = 0;
                sp_DataWord.ClearData();
                CF_DataWord.ClearData();
                EE_DataWord.ClearData();
                CIS_DataWord.ClearData();
                Noise_DataWord.ClearData();
                Absolute_RF_Power.ClearData();
                RF_Difference_DataWord.ClearData();
                RF_Gain_DataWord.ClearData();
                XML_Arrays.SP_ClearData();
                XML_Arrays.EE_ClearData();
                XML_Arrays.CF_ClearData();
                XML_Arrays.CIS_ClearData();
                XML_Arrays.Noise_ClearData();
                XML_Arrays.Absolute_RF_Power_ClearData();
                XML_Arrays.RF_Difference_ClearData();
                XML_Arrays.XML_RFG_ClearData();
                ExcelPage_ComboBox.Enabled = false;
                ReceiveData_Button.Enabled = false;
                SelectExcel_Button.Enabled = false;
            }


        }



        private void CreateCertificate_Button_Click(object sender, EventArgs e)
        {

            LabelProgress.Visible = false;
            progressBar.Value = 0;

            try
            {
                // Kullanıcıdan kaydedilecek dosya yolu al
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Title = "Save Certificates";
                saveFileDialog1.Filter = "Human and Machine Readable Certificate (.xml)(.docx)|*.docx";
                DialogResult result = saveFileDialog1.ShowDialog();

                // Kullanıcı bir konum seçtiyse devam et
                if (result == DialogResult.OK && !string.IsNullOrEmpty(saveFileDialog1.FileName))
                {
                    string originalFilePath = "wordData/sertifikaC.docx";
                    string copyFilePath = saveFileDialog1.FileName;

                    // Belgeyi kopyala
                    File.Copy(originalFilePath, copyFilePath, true);

                    // Word belgesine tablo ekle
                    createtable.ResultPages(tables, copyFilePath);

                    // XML dosyasını oluştur
                    string xmlFilePath = Path.ChangeExtension(copyFilePath, ".xml");
                    xml.Save(xmlFilePath);

                    // Bildirim mesajı göster
                    MessageBox.Show("Human-readable and machine-readable certificates were created.", "Successful !", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception err)
            {
                LabelProgress.Visible = true;
                LabelProgress.ForeColor = System.Drawing.Color.Red;
                LabelProgress.Text = @"ERROR!: Word";
                MessageBox.Show(err.Message, err.StackTrace, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            SelectExcel_Button.Enabled = false;
            CreateCertificate_Button.Enabled = false;
            LabelProgress.Text = "";
            progressBar.Value = 0;
            UncheckAllCheckBoxes(this);
            EnabledAllCheckBoxes(this);
            ExcelPage_ComboBox.Items.Clear();
            ExcelPage_ComboBox.Refresh();
            ExcelFileName_TextBox.Hint = "Please Select Xml File";
            ExcelFileName_TextBox.Text = "";
            listBox1.Items.Add("Certificate was created from the files above!");
            listBox1.Items.Add("------------------------------------------------------------------------");
            label7.Visible = false;
            xml.Load(xmlFilePath);
            tables.Clear();
            listboxnum =0;
        }
        #endregion

        #region ProgressBar Control 

        public void ImportDataBasarim()
        {
            LabelProgress.Text = @"Import data successfull";

            Thread.Sleep(10);
            progressBar.Value = 0;
            for (int i = 0; i < 100; i++)
            {
                progressBar.Value += 1;
            }
            LabelProgress.Visible = true;
            LabelProgress.ForeColor = System.Drawing.Color.Green;

        }
        public void KontrolBasarim()
        {
            // Kullanıcıya başarı mesajı göster
            MessageBox.Show("The table structure according to the selected cell is suitable for creating certificates. You can create a certificate.", "Successful!", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // label4'ün konumunu ayarla ve seçilen hücre bilgilerini göster
            label4.Location = new Point(396, 53);
            label4.Text = ($"Selection cell:  {"Column: "}{sütun}{"  Row: "}{satır}");

            // İlerleme durumunu gösteren etiketi güncelle
            LabelProgress.Text = "Cell selection successful !";

            // Kısa bir süre uyku (gecikme)
            Thread.Sleep(10);

            // ProgressBar'ı sıfırla ve doldur
            progressBar.Value = 0;
            for (int i = 0; i < 100; i++)
            {
                progressBar.Value += 1;
            }

            // İlerleme durumunu gösteren etiketi görünür yap ve rengini yeşil olarak ayarla
            LabelProgress.Visible = true;
            LabelProgress.ForeColor = System.Drawing.Color.Green;
        }


        #endregion

        #region SelectAll Button Kontrol


        private void CIS_SelectAll_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (CIS_SelectAll_CheckBox.Checked == true)
            {
                CIS_CheckBox.Checked = true;
                CIS_CheckBox1.Checked = true;
                CIS_CheckBox2.Checked = true;
                CIS_CheckBox3.Checked = true;
                CIS_CheckBox4.Checked = true;
                CIS_CheckBox5.Checked = true;
                CIS_CheckBox6.Checked = true;
            }
            if (CIS_SelectAll_CheckBox.Checked == false)
            {
                CIS_CheckBox.Checked = false;
                CIS_CheckBox1.Checked = false;
                CIS_CheckBox2.Checked = false;
                CIS_CheckBox3.Checked = false;
                CIS_CheckBox4.Checked = false;
                CIS_CheckBox5.Checked = false;
                CIS_CheckBox6.Checked = false;
            }
        }

        private void ARFP_SelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (ARFP_SelectAll.Checked == true)
            {
                ARFP_1.Checked = true;
                ARFP_2.Checked = true;
                ARFP_3.Checked = true;
                ARFP_4.Checked = true;
                ARFP_5.Checked = true;
                ARFP_6.Checked = true;
                ARFP_7.Checked = true;
                ARFP_8.Checked = true;
                ARFP_9.Checked = true;
                ARFP_10.Checked = true;
                ARFP_11.Checked = true;
            }
            if (ARFP_SelectAll.Checked == false)
            {
                ARFP_1.Checked = false;
                ARFP_2.Checked = false;
                ARFP_3.Checked = false;
                ARFP_4.Checked = false;
                ARFP_5.Checked = false;
                ARFP_6.Checked = false;
                ARFP_7.Checked = false;
                ARFP_8.Checked = false;
                ARFP_9.Checked = false;
                ARFP_10.Checked = false;
                ARFP_11.Checked = false;
            }
        }

        private void RF_Diff_SelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (RF_Diff_SelectAll.Checked == true)
            {
                RF_Diff_1.Checked = true;
                RF_Diff_2.Checked = true;
                RF_Diff_3.Checked = true;
                RF_Diff_4.Checked = true;
            }
            if (RF_Diff_SelectAll.Checked == false)
            {
                RF_Diff_1.Checked = false;
                RF_Diff_2.Checked = false;
                RF_Diff_3.Checked = false;
                RF_Diff_4.Checked = false;
            }
        }

        private void RF_Gain_SelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (RF_Gain_SelectAll.Checked == true)
            {
                RF_Gain1.Checked = true;
                RF_Gain2.Checked = true;
                RF_Gain3.Checked = true;
                RF_Gain4.Checked = true;

            }
            if (RF_Gain_SelectAll.Checked == false)
            {
                RF_Gain1.Checked = false;
                RF_Gain2.Checked = false;
                RF_Gain3.Checked = false;
                RF_Gain4.Checked = false;

            }
        }


        private void EE_SelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (EE_SelectAll.Checked == true)
            {
                checkBoxEE.Checked = true;
                checkBox_EE_RI.Checked = true;
                checkBoxRHO.Checked = true;
                checkBox_EE_CF.Checked = true;

            }
            if (EE_SelectAll.Checked == false)
            {
                checkBoxEE.Checked = false;
                checkBox_EE_RI.Checked = false;
                checkBoxRHO.Checked = false;
                checkBox_EE_CF.Checked = false;

            }
        }


        private void S_Parameter_SelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (S_Parameter_SelectAll.Checked == true)
            {
                checkBoxS11Reel.Checked = true;
                checkBoxS11Log.Checked = true;
                checkBoxS11Lin.Checked = true;
                checkBoxS11SWR.Checked = true;
                checkBoxS12Reel.Checked = true;
                checkBoxS12Log.Checked = true;
                checkBoxS12Lin.Checked = true;
                checkBoxS21Reel.Checked = true;
                checkBoxS21Log.Checked = true;
                checkBoxS21Lin.Checked = true;
                checkBoxS22Reel.Checked = true;
                checkBoxS22Log.Checked = true;
                checkBoxS22Lin.Checked = true;
                checkBoxS22SWR.Checked = true;
            }
            if (S_Parameter_SelectAll.Checked == false)
            {
                checkBoxS11Reel.Checked = false;
                checkBoxS11Log.Checked = false;
                checkBoxS11Lin.Checked = false;
                checkBoxS11SWR.Checked = false;
                checkBoxS12Reel.Checked = false;
                checkBoxS12Log.Checked = false;
                checkBoxS12Lin.Checked = false;
                checkBoxS21Reel.Checked = false;
                checkBoxS21Log.Checked = false;
                checkBoxS21Lin.Checked = false;
                checkBoxS22Reel.Checked = false;
                checkBoxS22Log.Checked = false;
                checkBoxS22Lin.Checked = false;
                checkBoxS22SWR.Checked = false;
            }

        }

        private void NSSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (NSSelectAll.Checked == true)
            {
                NS_checkBox_DC_ON.Checked = true;
                NS_checkBox_DC_OFF.Checked = true;
                NS_checkBoxENR.Checked = true;
            }
            if (NSSelectAll.Checked == false)
            {
                NS_checkBox_DC_ON.Checked = false;
                NS_checkBox_DC_OFF.Checked = false;
                NS_checkBoxENR.Checked = false;
            }
        }

        private void CFSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (CFSelectAll.Checked == true)
            {
                CheckBox_CF.Checked = true;
                CF_checkBox_RIRC.Checked = true;
            }
            if (CFSelectAll.Checked == false)
            {
                CheckBox_CF.Checked = false;
                CF_checkBox_RIRC.Checked = false;
            }

        }
        #endregion

        #region Checkbox Seçimine göre kısıtlama kontrolleri

        // CheckBoxTabControl içindeki tüm CheckBox'ların CheckedChanged olayına event handler ekler
        public void CheckBoxTabpagecontrol()
        {
            // CheckBoxTabControl içindeki tüm kontrolleri dolaş
            foreach (Control control in CheckBoxTabControl.Controls)
            {
                // Eğer kontrol bir TabPage ise
                if (control is TabPage tabPage)
                {
                    // TabPage içindeki tüm kontrolleri dolaş
                    foreach (Control tabPageControl in tabPage.Controls)
                    {
                        // Eğer kontrol bir CheckBox ise
                        if (tabPageControl is CheckBox checkBox)
                        {
                            // CheckBox'ın CheckedChanged olayına event handler ekle
                            checkBox.CheckedChanged += CheckBox_CheckedChanged;
                        }
                    }
                }
            }
        }

        // RFPowTabControl içindeki tüm CheckBox'ların CheckedChanged olayına event handler ekler
        public void RFPowtabpageControl()
        {
            // RFPowTabControl içindeki tüm kontrolleri dolaş
            foreach (Control control in RFPowTabControl.Controls)
            {
                // Eğer kontrol bir TabPage ise
                if (control is TabPage tabPage)
                {
                    // TabPage içindeki tüm kontrolleri dolaş
                    foreach (Control tabPageControl in tabPage.Controls)
                    {
                        // Eğer kontrol bir CheckBox ise
                        if (tabPageControl is CheckBox checkBox)
                        {
                            // CheckBox'ın CheckedChanged olayına event handler ekle
                            checkBox.CheckedChanged += RFPowtabpageCheckBox_CheckedChanged;
                        }
                    }
                }
            }
        }

    


        private void RFPowtabpageCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            // CheckBox'ın durumuna göre ComboBox'ın tıklanabilirliğini ayarla
            bool anyCheckBoxChecked = RFPowTabControl.TabPages.Cast<TabPage>()
                .SelectMany(tabPage => tabPage.Controls.Cast<Control>()
                    .Where(control => control is CheckBox)
                    .Cast<CheckBox>())
                .Any(checkBox => checkBox.Checked);

            MeasurementTypes_ComboBox.Enabled = !anyCheckBoxChecked;
            materialTabSelector1.Enabled = !anyCheckBoxChecked;
            SelectExcel_Button.Enabled = anyCheckBoxChecked;
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            // CheckBox'ın durumuna göre ComboBox'ın tıklanabilirliğini ayarla
            bool anyCheckBoxChecked = CheckBoxTabControl.TabPages.Cast<TabPage>()
                .SelectMany(tabPage => tabPage.Controls.Cast<Control>()
                    .Where(control => control is CheckBox)
                    .Cast<CheckBox>())
                .Any(checkBox => checkBox.Checked);

            MeasurementTypes_ComboBox.Enabled = !anyCheckBoxChecked;
            SelectExcel_Button.Enabled = anyCheckBoxChecked;

        }

        private void UncheckAllCheckBoxes(Control container)
        {
            // Konteyner içindeki tüm kontrol elemanlarını döngü ile kontrol et
            foreach (Control control in container.Controls)
            {
                // Eğer kontrol elemanı bir konteyner ise (groupbox, panel vb.), içindeki kontrol elemanlarını kontrol et
                if (control.HasChildren)
                {
                    UncheckAllCheckBoxes(control);
                }

                // Kontrol elemanı bir checkbox ise, checked özelliğini false yap
                if (control is CheckBox)
                {
                    CheckBox checkBox = (CheckBox)control;
                    checkBox.Checked = false;
                }
            }
        }

        private void UnEnabledAllCheckBoxes(Control container)
        {
            // Konteyner içindeki tüm kontrol elemanlarını döngü ile kontrol et
            foreach (Control control in container.Controls)
            {
                // Eğer kontrol elemanı bir konteyner ise (groupbox, panel vb.), içindeki kontrol elemanlarını kontrol et
                if (control.HasChildren)
                {
                    UnEnabledAllCheckBoxes(control);
                }

                // Kontrol elemanı bir checkbox ise, checked özelliğini false yap
                if (control is CheckBox)
                {
                    CheckBox checkBox = (CheckBox)control;
                    checkBox.Enabled = false;
                }
            }
        }

        private void EnabledAllCheckBoxes(Control container)
        {
            // Konteyner içindeki tüm kontrol elemanlarını döngü ile kontrol et
            foreach (Control control in container.Controls)
            {
                // Eğer kontrol elemanı bir konteyner ise (groupbox, panel vb.), içindeki kontrol elemanlarını kontrol et
                if (control.HasChildren)
                {
                    EnabledAllCheckBoxes(control);
                }

                // Kontrol elemanı bir checkbox ise, checked özelliğini false yap
                if (control is CheckBox)
                {
                    CheckBox checkBox = (CheckBox)control;
                    checkBox.Enabled = true;
                }
            }
        }

        #endregion

        #region Excel Şablon kontrolü ve Datagridview kodları 

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
                        try
                        {
                            object value = worksheet.Cells[i, j].Value;
                            row.Cells[j - 1].Value = value != null ? value.ToString() : "";
                        }
                        catch (Exception ex)
                        {
                            // Hataları yakala ve hücreyi boş bırak
                            row.Cells[j - 1].Value = "";
                        }
                    }

                    dataGridView1.Rows.Add(row);
                }

                // DataGridView'e sütun başlıklarını gösterme
                dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
        }
        private string GetExcelColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = string.Empty;

            // Excel sütun isimlerini belirlemek için döngü
            while (dividend > 0)
            {
                // Modulus (kalan) hesapla
                int modulo = (dividend - 1) % 26;
                // Modulus değerine göre karakteri belirle (A = 65 ASCII değerine sahip)
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                // Bölümü yeniden hesapla
                dividend = (dividend - modulo) / 26;
            }

            return columnName; // Hesaplanan sütun ismini döndür
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // Farklı ölçüm tipleri için boolean listeleri
            List<bool> sparambool = new List<bool>() { true, true, true, true, true, false, true, true, true, true, false, true, true, true, true, false,
                                               true, true, false, true, true, true, true, false, true, true, true, true, false, true, true,
                                               true, true, false, true, true, true, true, false, true, true, true, true, false, true, true,
                                               true, true, false, true, true, true, true, false, true, true, true, true, false, true, true,
                                               true, true, false, true, true };
            List<bool> EEbool = new List<bool>() { true, true, true, false, true, true, true, true, false, true, true, false, true, true };
            List<bool> CFbool = new List<bool>() { true, true, true, false, true, true, true, true, true, true, true };
            List<bool> CISbool = new List<bool>() { true, true, true, true, true, true, true };
            List<bool> ABSbool = new List<bool>() { true, true, true, true, true, true, true, false, true, true, true, true, true, true, true, false, true, true, true, true,
                                            true, true, true, false, true, true, true, true, true, false, true, true, true, true, true, false, true, true, true, true,
                                            true, false, true, true, true, true, true, true, true, false, true, true, true, true, true, true, true, false, true, true,
                                            true, true, true, false, true, true, true, true, true };
            List<bool> RFDbool = new List<bool>() { true, true, true, true, true, true, true, false, true, true, true, true, true, true, true, false, true, true, true, true,
                                            true, true, true, false, true, true, true, true, true, true, true };
            List<bool> RFGbool = new List<bool>() { true, true, true, false, true, true, true, false, true, true, true, false, true, true, true };
            List<bool> NSbool = new List<bool>() { true, true, true, false, true, true, true, true, true, true, false, true, true, true, true, true, true };

            // Eğer geçerli bir hücreye çift tıklanmışsa
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewCell selectedCell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                string selectedWorksheet = ExcelPage_ComboBox.SelectedItem.ToString();
                FileInfo fileInfo = new FileInfo(openFileDialog.FileName);

                using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[selectedWorksheet];
                    string selectedColumnHeader = selectedCell.OwningColumn.HeaderText; // Seçilen sütunun başlığını al
                    int rowIndex = e.RowIndex;
                    satır = rowIndex + 1;
                    sütun = selectedColumnHeader;

                    int startRow = worksheet.Dimension.Start.Row; // Tablonun başlangıç satırı
                    int endRow = worksheet.Dimension.End.Row;     // Tablonun son satırı
                    int lastColumnWithData = worksheet.Dimension.End.Column; // Veri bulunan son sütun

                    List<bool> boolList = new List<bool>();

                    // Kullanıcının seçtiği sütundan başlayarak veri bulunan son sütuna kadar tara
                    for (int col = selectedCell.ColumnIndex + 1; col <= lastColumnWithData; col++)
                    {
                        bool hasData = false;

                        // Herhangi bir veri içeren bir hücre var mı kontrol et
                        for (int row = startRow; row <= endRow; row++)
                        {
                            if (!string.IsNullOrEmpty(worksheet.Cells[row, col].Text))
                            {
                                Save_Row_Col_Button.Enabled = true;
                                hasData = true;
                                break;
                            }
                        }

                        // Bool listesine sonucu ekle
                        boolList.Add(hasData);
                    }

                    // Seçilen ölçüm tipine göre bool listeleri ile karşılaştırma yap
                    if (MeasurementTypes_ComboBox.SelectedIndex == 0 && AreListsEqual(boolList, EEbool))
                    {
                        KontrolBasarim();
                    }
                    else if (MeasurementTypes_ComboBox.SelectedIndex == 1 && AreListsEqual(boolList, CFbool))
                    {
                        KontrolBasarim();
                    }
                    else if (MeasurementTypes_ComboBox.SelectedIndex == 2 && AreListsEqual(boolList, CISbool))
                    {
                        KontrolBasarim();
                    }
                    else if (MeasurementTypes_ComboBox.SelectedIndex == 3 && RFPowTabControl.SelectedTab == Abs_RF_Power_tabpage && AreListsEqual(boolList, ABSbool))
                    {
                        KontrolBasarim();
                    }
                    else if (MeasurementTypes_ComboBox.SelectedIndex == 3 && RFPowTabControl.SelectedTab == RF_Diff_tabpage && AreListsEqual(boolList, RFDbool))
                    {
                        KontrolBasarim();
                    }
                    else if (MeasurementTypes_ComboBox.SelectedIndex == 3 && RFPowTabControl.SelectedTab == RF_Gain_tabpage && AreListsEqual(boolList, RFGbool))
                    {
                        KontrolBasarim();
                    }
                    else if (MeasurementTypes_ComboBox.SelectedIndex == 4 && AreListsEqual(boolList, sparambool))
                    {
                        KontrolBasarim();
                    }
                    else if (MeasurementTypes_ComboBox.SelectedIndex == 5 && AreListsEqual(boolList, NSbool))
                    {
                        KontrolBasarim();
                    }
                    else
                    {
                        // Uygun olmayan tablo yapısı için hata mesajı göster
                        MessageBox.Show("The table structure is not compatible with the selected parameter, please check the cell you selected and the Excel file you selected!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Save_Row_Col_Button.Enabled = false;
                        label4.Text = ($"Selection cell:  {"Column: "}{sütun}{"  Row: "}{satır}");
                        LabelProgress.Text = "";
                        progressBar.Value = 0;
                    }
                }
            }
        }


        static bool AreListsEqual(List<bool> list1, List<bool> list2)
        {
            // İki listenin uzunlukları eşit değilse, listeler eşit değildir
            if (list1.Count != list2.Count)
                return false;

            // Her iki listedeki elemanları karşılaştır
            for (int i = 0; i < list1.Count; i++)
            {
                // Herhangi bir eleman farklıysa, listeler eşit değildir
                if (list1[i] != list2[i])
                    return false;
            }

            // Tüm elemanlar aynı ise, listeler eşittir
            return true;
        }


        #endregion

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}

