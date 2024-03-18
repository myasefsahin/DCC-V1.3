namespace DCC
{
    partial class CertificateForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CertificateForm));
            BackBox1 = new PictureBox();
            CertificateTabControl = new TabControl();
            API_PAGE = new TabPage();
            panel1 = new Panel();
            groupBox4 = new GroupBox();
            NextButton = new MaterialSkin.Controls.MaterialButton();
            MeasurementsTextBox = new MaterialSkin.Controls.MaterialMultiLineTextBox2();
            CalibrationDescTextBox = new MaterialSkin.Controls.MaterialMultiLineTextBox2();
            MethodTextBox = new MaterialSkin.Controls.MaterialMultiLineTextBox2();
            DeviceTextBox = new MaterialSkin.Controls.MaterialMultiLineTextBox2();
            label2 = new Label();
            groupBox1 = new GroupBox();
            SelectDeviceButton = new MaterialSkin.Controls.MaterialButton();
            groupBox3 = new GroupBox();
            CalCodeTextBox = new MaterialSkin.Controls.MaterialTextBox2();
            SerialNumberTextBox = new MaterialSkin.Controls.MaterialTextBox2();
            ModelNameTextBox = new MaterialSkin.Controls.MaterialTextBox2();
            DeviceNameTextBox = new MaterialSkin.Controls.MaterialTextBox2();
            groupBox2 = new GroupBox();
            LaboratoryComboBox = new MaterialSkin.Controls.MaterialComboBox();
            OrderNumberTextBox = new MaterialSkin.Controls.MaterialTextBox2();
            label1 = new Label();
            DATA_PAGE = new TabPage();
            panel2 = new Panel();
            groupBox10 = new GroupBox();
            CreateCertificate_Button = new MaterialSkin.Controls.MaterialButton();
            groupBox11 = new GroupBox();
            ReceiveData_Button = new MaterialSkin.Controls.MaterialButton();
            SelectExcel_Button = new MaterialSkin.Controls.MaterialButton();
            ExcelPage_ComboBox = new MaterialSkin.Controls.MaterialComboBox();
            ExcelFileName_TextBox = new MaterialSkin.Controls.MaterialTextBox2();
            groupBox5 = new GroupBox();
            CheckBoxTabControl = new MaterialSkin.Controls.MaterialTabControl();
            EE_Page = new TabPage();
            checkBox_EE_CF = new CheckBox();
            checkBoxRHO = new CheckBox();
            checkBox_EE_RI = new CheckBox();
            checkBoxEE = new CheckBox();
            CalFactor_Page = new TabPage();
            CF_checkBox_RIRC = new CheckBox();
            CheckBox_CF = new CheckBox();
            CIS_Page = new TabPage();
            CIS_CheckBox = new CheckBox();
            RFPow_Page = new TabPage();
            SParam_Page = new TabPage();
            groupBox9 = new GroupBox();
            checkBoxS22SWR = new CheckBox();
            checkBoxS22Log = new CheckBox();
            checkBoxS22Lin = new CheckBox();
            checkBoxS22Reel = new CheckBox();
            groupBox8 = new GroupBox();
            checkBoxS21Log = new CheckBox();
            checkBoxS21Lin = new CheckBox();
            checkBoxS21Reel = new CheckBox();
            groupBox7 = new GroupBox();
            checkBoxS12Log = new CheckBox();
            checkBoxS12Lin = new CheckBox();
            checkBoxS12Reel = new CheckBox();
            groupBox6 = new GroupBox();
            checkBoxS11SWR = new CheckBox();
            checkBoxS11Log = new CheckBox();
            checkBoxS11Lin = new CheckBox();
            checkBoxS11Reel = new CheckBox();
            MetCH_Page = new TabPage();
            Noise_Page = new TabPage();
            MeasurementTypes_ComboBox = new MaterialSkin.Controls.MaterialComboBox();
            label3 = new Label();
            BackBox2 = new PictureBox();
            ExcelView_Page = new TabPage();
            panel3 = new Panel();
            Save_Row_Col_Button = new MaterialSkin.Controls.MaterialButton();
            RowNumberTextBox = new MaterialSkin.Controls.MaterialTextBox2();
            ColumnNameTextBox = new MaterialSkin.Controls.MaterialTextBox2();
            dataGridView1 = new DataGridView();
            label4 = new Label();
            BackBox3 = new PictureBox();
            progressBar = new ProgressBar();
            label6 = new Label();
            groupBox12 = new GroupBox();
            LabelProgress = new Label();
            ((System.ComponentModel.ISupportInitialize)BackBox1).BeginInit();
            CertificateTabControl.SuspendLayout();
            API_PAGE.SuspendLayout();
            panel1.SuspendLayout();
            groupBox4.SuspendLayout();
            groupBox1.SuspendLayout();
            groupBox3.SuspendLayout();
            groupBox2.SuspendLayout();
            DATA_PAGE.SuspendLayout();
            panel2.SuspendLayout();
            groupBox10.SuspendLayout();
            groupBox11.SuspendLayout();
            groupBox5.SuspendLayout();
            CheckBoxTabControl.SuspendLayout();
            EE_Page.SuspendLayout();
            CalFactor_Page.SuspendLayout();
            CIS_Page.SuspendLayout();
            SParam_Page.SuspendLayout();
            groupBox9.SuspendLayout();
            groupBox8.SuspendLayout();
            groupBox7.SuspendLayout();
            groupBox6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)BackBox2).BeginInit();
            ExcelView_Page.SuspendLayout();
            panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)BackBox3).BeginInit();
            groupBox12.SuspendLayout();
            SuspendLayout();
            // 
            // BackBox1
            // 
            BackBox1.Image = (Image)resources.GetObject("BackBox1.Image");
            BackBox1.Location = new Point(5, 14);
            BackBox1.Name = "BackBox1";
            BackBox1.Size = new Size(62, 35);
            BackBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            BackBox1.TabIndex = 6;
            BackBox1.TabStop = false;
            BackBox1.Click += BackBox1_Click;
            // 
            // CertificateTabControl
            // 
            CertificateTabControl.Controls.Add(API_PAGE);
            CertificateTabControl.Controls.Add(DATA_PAGE);
            CertificateTabControl.Controls.Add(ExcelView_Page);
            CertificateTabControl.Location = new Point(0, -34);
            CertificateTabControl.Name = "CertificateTabControl";
            CertificateTabControl.SelectedIndex = 0;
            CertificateTabControl.Size = new Size(1227, 754);
            CertificateTabControl.TabIndex = 7;
            // 
            // API_PAGE
            // 
            API_PAGE.Controls.Add(panel1);
            API_PAGE.Location = new Point(4, 29);
            API_PAGE.Name = "API_PAGE";
            API_PAGE.Padding = new Padding(3);
            API_PAGE.Size = new Size(1219, 721);
            API_PAGE.TabIndex = 0;
            API_PAGE.Text = "tabPage1";
            API_PAGE.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            panel1.BackColor = Color.White;
            panel1.Controls.Add(groupBox4);
            panel1.Controls.Add(groupBox1);
            panel1.Controls.Add(BackBox1);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(3, 3);
            panel1.Name = "panel1";
            panel1.Size = new Size(1213, 715);
            panel1.TabIndex = 7;
            // 
            // groupBox4
            // 
            groupBox4.BackColor = Color.White;
            groupBox4.Controls.Add(NextButton);
            groupBox4.Controls.Add(MeasurementsTextBox);
            groupBox4.Controls.Add(CalibrationDescTextBox);
            groupBox4.Controls.Add(MethodTextBox);
            groupBox4.Controls.Add(DeviceTextBox);
            groupBox4.Controls.Add(label2);
            groupBox4.Location = new Point(540, 3);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new Size(664, 721);
            groupBox4.TabIndex = 8;
            groupBox4.TabStop = false;
            // 
            // NextButton
            // 
            NextButton.AutoSize = false;
            NextButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            NextButton.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            NextButton.Depth = 0;
            NextButton.HighEmphasis = true;
            NextButton.Icon = null;
            NextButton.Location = new Point(470, 660);
            NextButton.Margin = new Padding(4, 6, 4, 6);
            NextButton.MouseState = MaterialSkin.MouseState.HOVER;
            NextButton.Name = "NextButton";
            NextButton.NoAccentTextColor = Color.Empty;
            NextButton.Size = new Size(127, 36);
            NextButton.TabIndex = 6;
            NextButton.Text = "save and next";
            NextButton.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            NextButton.UseAccentColor = false;
            NextButton.UseVisualStyleBackColor = true;
            NextButton.Click += NextButton_Click;
            // 
            // MeasurementsTextBox
            // 
            MeasurementsTextBox.AnimateReadOnly = false;
            MeasurementsTextBox.BackgroundImageLayout = ImageLayout.None;
            MeasurementsTextBox.CharacterCasing = CharacterCasing.Normal;
            MeasurementsTextBox.Depth = 0;
            MeasurementsTextBox.HideSelection = true;
            MeasurementsTextBox.Hint = "MEASUREMENTS";
            MeasurementsTextBox.Location = new Point(76, 505);
            MeasurementsTextBox.MaxLength = 32767;
            MeasurementsTextBox.MouseState = MaterialSkin.MouseState.OUT;
            MeasurementsTextBox.Name = "MeasurementsTextBox";
            MeasurementsTextBox.PasswordChar = '\0';
            MeasurementsTextBox.ReadOnly = false;
            MeasurementsTextBox.ScrollBars = ScrollBars.None;
            MeasurementsTextBox.SelectedText = "";
            MeasurementsTextBox.SelectionLength = 0;
            MeasurementsTextBox.SelectionStart = 0;
            MeasurementsTextBox.ShortcutsEnabled = true;
            MeasurementsTextBox.Size = new Size(512, 101);
            MeasurementsTextBox.TabIndex = 5;
            MeasurementsTextBox.TabStop = false;
            MeasurementsTextBox.TextAlign = HorizontalAlignment.Left;
            MeasurementsTextBox.UseSystemPasswordChar = false;
            // 
            // CalibrationDescTextBox
            // 
            CalibrationDescTextBox.AnimateReadOnly = false;
            CalibrationDescTextBox.BackgroundImageLayout = ImageLayout.None;
            CalibrationDescTextBox.CharacterCasing = CharacterCasing.Normal;
            CalibrationDescTextBox.Depth = 0;
            CalibrationDescTextBox.HideSelection = true;
            CalibrationDescTextBox.Hint = "CALİBRATİON";
            CalibrationDescTextBox.Location = new Point(76, 362);
            CalibrationDescTextBox.MaxLength = 32767;
            CalibrationDescTextBox.MouseState = MaterialSkin.MouseState.OUT;
            CalibrationDescTextBox.Name = "CalibrationDescTextBox";
            CalibrationDescTextBox.PasswordChar = '\0';
            CalibrationDescTextBox.ReadOnly = false;
            CalibrationDescTextBox.ScrollBars = ScrollBars.None;
            CalibrationDescTextBox.SelectedText = "";
            CalibrationDescTextBox.SelectionLength = 0;
            CalibrationDescTextBox.SelectionStart = 0;
            CalibrationDescTextBox.ShortcutsEnabled = true;
            CalibrationDescTextBox.Size = new Size(512, 101);
            CalibrationDescTextBox.TabIndex = 4;
            CalibrationDescTextBox.TabStop = false;
            CalibrationDescTextBox.TextAlign = HorizontalAlignment.Left;
            CalibrationDescTextBox.UseSystemPasswordChar = false;
            // 
            // MethodTextBox
            // 
            MethodTextBox.AnimateReadOnly = false;
            MethodTextBox.BackgroundImageLayout = ImageLayout.None;
            MethodTextBox.CharacterCasing = CharacterCasing.Normal;
            MethodTextBox.Depth = 0;
            MethodTextBox.HideSelection = true;
            MethodTextBox.Hint = "METHOD";
            MethodTextBox.Location = new Point(76, 234);
            MethodTextBox.MaxLength = 32767;
            MethodTextBox.MouseState = MaterialSkin.MouseState.OUT;
            MethodTextBox.Name = "MethodTextBox";
            MethodTextBox.PasswordChar = '\0';
            MethodTextBox.ReadOnly = false;
            MethodTextBox.ScrollBars = ScrollBars.None;
            MethodTextBox.SelectedText = "";
            MethodTextBox.SelectionLength = 0;
            MethodTextBox.SelectionStart = 0;
            MethodTextBox.ShortcutsEnabled = true;
            MethodTextBox.Size = new Size(512, 101);
            MethodTextBox.TabIndex = 3;
            MethodTextBox.TabStop = false;
            MethodTextBox.TextAlign = HorizontalAlignment.Left;
            MethodTextBox.UseSystemPasswordChar = false;
            // 
            // DeviceTextBox
            // 
            DeviceTextBox.AnimateReadOnly = false;
            DeviceTextBox.BackgroundImageLayout = ImageLayout.None;
            DeviceTextBox.CharacterCasing = CharacterCasing.Normal;
            DeviceTextBox.Depth = 0;
            DeviceTextBox.HideSelection = true;
            DeviceTextBox.Hint = "DEVİCE";
            DeviceTextBox.Location = new Point(76, 103);
            DeviceTextBox.MaxLength = 32767;
            DeviceTextBox.MouseState = MaterialSkin.MouseState.OUT;
            DeviceTextBox.Name = "DeviceTextBox";
            DeviceTextBox.PasswordChar = '\0';
            DeviceTextBox.ReadOnly = false;
            DeviceTextBox.ScrollBars = ScrollBars.None;
            DeviceTextBox.SelectedText = "";
            DeviceTextBox.SelectionLength = 0;
            DeviceTextBox.SelectionStart = 0;
            DeviceTextBox.ShortcutsEnabled = true;
            DeviceTextBox.Size = new Size(512, 101);
            DeviceTextBox.TabIndex = 2;
            DeviceTextBox.TabStop = false;
            DeviceTextBox.TextAlign = HorizontalAlignment.Left;
            DeviceTextBox.UseSystemPasswordChar = false;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point, 162);
            label2.ForeColor = Color.Navy;
            label2.Location = new Point(52, 32);
            label2.Name = "label2";
            label2.Size = new Size(560, 23);
            label2.TabIndex = 0;
            label2.Text = "Device, Method, Calibration and Measurement Description Windows";
            // 
            // groupBox1
            // 
            groupBox1.BackColor = Color.White;
            groupBox1.Controls.Add(SelectDeviceButton);
            groupBox1.Controls.Add(groupBox3);
            groupBox1.Controls.Add(groupBox2);
            groupBox1.Controls.Add(label1);
            groupBox1.Location = new Point(92, 3);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(433, 721);
            groupBox1.TabIndex = 7;
            groupBox1.TabStop = false;
            // 
            // SelectDeviceButton
            // 
            SelectDeviceButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            SelectDeviceButton.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            SelectDeviceButton.Depth = 0;
            SelectDeviceButton.HighEmphasis = true;
            SelectDeviceButton.Icon = null;
            SelectDeviceButton.Location = new Point(138, 660);
            SelectDeviceButton.Margin = new Padding(4, 6, 4, 6);
            SelectDeviceButton.MouseState = MaterialSkin.MouseState.HOVER;
            SelectDeviceButton.Name = "SelectDeviceButton";
            SelectDeviceButton.NoAccentTextColor = Color.Empty;
            SelectDeviceButton.Size = new Size(127, 36);
            SelectDeviceButton.TabIndex = 3;
            SelectDeviceButton.Text = "Select Device";
            SelectDeviceButton.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            SelectDeviceButton.UseAccentColor = false;
            SelectDeviceButton.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            groupBox3.BackColor = Color.White;
            groupBox3.Controls.Add(CalCodeTextBox);
            groupBox3.Controls.Add(SerialNumberTextBox);
            groupBox3.Controls.Add(ModelNameTextBox);
            groupBox3.Controls.Add(DeviceNameTextBox);
            groupBox3.Location = new Point(24, 275);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(385, 366);
            groupBox3.TabIndex = 2;
            groupBox3.TabStop = false;
            groupBox3.Text = "Device";
            // 
            // CalCodeTextBox
            // 
            CalCodeTextBox.AnimateReadOnly = false;
            CalCodeTextBox.BackgroundImageLayout = ImageLayout.None;
            CalCodeTextBox.CharacterCasing = CharacterCasing.Normal;
            CalCodeTextBox.Depth = 0;
            CalCodeTextBox.Font = new Font("Microsoft Sans Serif", 16F, FontStyle.Regular, GraphicsUnit.Pixel);
            CalCodeTextBox.HideSelection = true;
            CalCodeTextBox.Hint = "Please Enter Calibration Code";
            CalCodeTextBox.LeadingIcon = null;
            CalCodeTextBox.Location = new Point(37, 283);
            CalCodeTextBox.MaxLength = 32767;
            CalCodeTextBox.MouseState = MaterialSkin.MouseState.OUT;
            CalCodeTextBox.Name = "CalCodeTextBox";
            CalCodeTextBox.PasswordChar = '\0';
            CalCodeTextBox.PrefixSuffixText = null;
            CalCodeTextBox.ReadOnly = false;
            CalCodeTextBox.RightToLeft = RightToLeft.No;
            CalCodeTextBox.SelectedText = "";
            CalCodeTextBox.SelectionLength = 0;
            CalCodeTextBox.SelectionStart = 0;
            CalCodeTextBox.ShortcutsEnabled = true;
            CalCodeTextBox.Size = new Size(312, 48);
            CalCodeTextBox.TabIndex = 3;
            CalCodeTextBox.TabStop = false;
            CalCodeTextBox.TextAlign = HorizontalAlignment.Left;
            CalCodeTextBox.TrailingIcon = null;
            CalCodeTextBox.UseSystemPasswordChar = false;
            // 
            // SerialNumberTextBox
            // 
            SerialNumberTextBox.AnimateReadOnly = false;
            SerialNumberTextBox.BackgroundImageLayout = ImageLayout.None;
            SerialNumberTextBox.CharacterCasing = CharacterCasing.Normal;
            SerialNumberTextBox.Depth = 0;
            SerialNumberTextBox.Font = new Font("Microsoft Sans Serif", 16F, FontStyle.Regular, GraphicsUnit.Pixel);
            SerialNumberTextBox.HideSelection = true;
            SerialNumberTextBox.Hint = "Please Enter Serial Number";
            SerialNumberTextBox.LeadingIcon = null;
            SerialNumberTextBox.Location = new Point(37, 204);
            SerialNumberTextBox.MaxLength = 32767;
            SerialNumberTextBox.MouseState = MaterialSkin.MouseState.OUT;
            SerialNumberTextBox.Name = "SerialNumberTextBox";
            SerialNumberTextBox.PasswordChar = '\0';
            SerialNumberTextBox.PrefixSuffixText = null;
            SerialNumberTextBox.ReadOnly = false;
            SerialNumberTextBox.RightToLeft = RightToLeft.No;
            SerialNumberTextBox.SelectedText = "";
            SerialNumberTextBox.SelectionLength = 0;
            SerialNumberTextBox.SelectionStart = 0;
            SerialNumberTextBox.ShortcutsEnabled = true;
            SerialNumberTextBox.Size = new Size(312, 48);
            SerialNumberTextBox.TabIndex = 2;
            SerialNumberTextBox.TabStop = false;
            SerialNumberTextBox.TextAlign = HorizontalAlignment.Left;
            SerialNumberTextBox.TrailingIcon = null;
            SerialNumberTextBox.UseSystemPasswordChar = false;
            // 
            // ModelNameTextBox
            // 
            ModelNameTextBox.AnimateReadOnly = false;
            ModelNameTextBox.BackgroundImageLayout = ImageLayout.None;
            ModelNameTextBox.CharacterCasing = CharacterCasing.Normal;
            ModelNameTextBox.Depth = 0;
            ModelNameTextBox.Font = new Font("Microsoft Sans Serif", 16F, FontStyle.Regular, GraphicsUnit.Pixel);
            ModelNameTextBox.HideSelection = true;
            ModelNameTextBox.Hint = "Please Enter Model Name";
            ModelNameTextBox.LeadingIcon = null;
            ModelNameTextBox.Location = new Point(37, 122);
            ModelNameTextBox.MaxLength = 32767;
            ModelNameTextBox.MouseState = MaterialSkin.MouseState.OUT;
            ModelNameTextBox.Name = "ModelNameTextBox";
            ModelNameTextBox.PasswordChar = '\0';
            ModelNameTextBox.PrefixSuffixText = null;
            ModelNameTextBox.ReadOnly = false;
            ModelNameTextBox.RightToLeft = RightToLeft.No;
            ModelNameTextBox.SelectedText = "";
            ModelNameTextBox.SelectionLength = 0;
            ModelNameTextBox.SelectionStart = 0;
            ModelNameTextBox.ShortcutsEnabled = true;
            ModelNameTextBox.Size = new Size(312, 48);
            ModelNameTextBox.TabIndex = 1;
            ModelNameTextBox.TabStop = false;
            ModelNameTextBox.TextAlign = HorizontalAlignment.Left;
            ModelNameTextBox.TrailingIcon = null;
            ModelNameTextBox.UseSystemPasswordChar = false;
            // 
            // DeviceNameTextBox
            // 
            DeviceNameTextBox.AnimateReadOnly = false;
            DeviceNameTextBox.BackgroundImageLayout = ImageLayout.None;
            DeviceNameTextBox.CharacterCasing = CharacterCasing.Normal;
            DeviceNameTextBox.Depth = 0;
            DeviceNameTextBox.Font = new Font("Microsoft Sans Serif", 16F, FontStyle.Regular, GraphicsUnit.Pixel);
            DeviceNameTextBox.HideSelection = true;
            DeviceNameTextBox.Hint = "Please Enter Device Name";
            DeviceNameTextBox.LeadingIcon = null;
            DeviceNameTextBox.Location = new Point(37, 41);
            DeviceNameTextBox.MaxLength = 32767;
            DeviceNameTextBox.MouseState = MaterialSkin.MouseState.OUT;
            DeviceNameTextBox.Name = "DeviceNameTextBox";
            DeviceNameTextBox.PasswordChar = '\0';
            DeviceNameTextBox.PrefixSuffixText = null;
            DeviceNameTextBox.ReadOnly = false;
            DeviceNameTextBox.RightToLeft = RightToLeft.No;
            DeviceNameTextBox.SelectedText = "";
            DeviceNameTextBox.SelectionLength = 0;
            DeviceNameTextBox.SelectionStart = 0;
            DeviceNameTextBox.ShortcutsEnabled = true;
            DeviceNameTextBox.Size = new Size(312, 48);
            DeviceNameTextBox.TabIndex = 0;
            DeviceNameTextBox.TabStop = false;
            DeviceNameTextBox.TextAlign = HorizontalAlignment.Left;
            DeviceNameTextBox.TrailingIcon = null;
            DeviceNameTextBox.UseSystemPasswordChar = false;
            // 
            // groupBox2
            // 
            groupBox2.BackColor = Color.White;
            groupBox2.Controls.Add(LaboratoryComboBox);
            groupBox2.Controls.Add(OrderNumberTextBox);
            groupBox2.Location = new Point(24, 71);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(385, 175);
            groupBox2.TabIndex = 1;
            groupBox2.TabStop = false;
            // 
            // LaboratoryComboBox
            // 
            LaboratoryComboBox.AutoResize = false;
            LaboratoryComboBox.BackColor = Color.FromArgb(255, 255, 255);
            LaboratoryComboBox.Depth = 0;
            LaboratoryComboBox.DrawMode = DrawMode.OwnerDrawVariable;
            LaboratoryComboBox.DropDownHeight = 174;
            LaboratoryComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            LaboratoryComboBox.DropDownWidth = 121;
            LaboratoryComboBox.Font = new Font("Microsoft Sans Serif", 14F, FontStyle.Bold, GraphicsUnit.Pixel);
            LaboratoryComboBox.ForeColor = Color.FromArgb(222, 0, 0, 0);
            LaboratoryComboBox.FormattingEnabled = true;
            LaboratoryComboBox.Hint = "Please Select Laboratory";
            LaboratoryComboBox.IntegralHeight = false;
            LaboratoryComboBox.ItemHeight = 43;
            LaboratoryComboBox.Location = new Point(37, 108);
            LaboratoryComboBox.MaxDropDownItems = 4;
            LaboratoryComboBox.MouseState = MaterialSkin.MouseState.OUT;
            LaboratoryComboBox.Name = "LaboratoryComboBox";
            LaboratoryComboBox.Size = new Size(312, 49);
            LaboratoryComboBox.StartIndex = 0;
            LaboratoryComboBox.TabIndex = 1;
            // 
            // OrderNumberTextBox
            // 
            OrderNumberTextBox.AnimateReadOnly = false;
            OrderNumberTextBox.BackgroundImageLayout = ImageLayout.None;
            OrderNumberTextBox.CharacterCasing = CharacterCasing.Normal;
            OrderNumberTextBox.Depth = 0;
            OrderNumberTextBox.Font = new Font("Microsoft Sans Serif", 16F, FontStyle.Regular, GraphicsUnit.Pixel);
            OrderNumberTextBox.HideSelection = true;
            OrderNumberTextBox.Hint = "Please Enter Order Number";
            OrderNumberTextBox.LeadingIcon = null;
            OrderNumberTextBox.Location = new Point(37, 26);
            OrderNumberTextBox.MaxLength = 32767;
            OrderNumberTextBox.MouseState = MaterialSkin.MouseState.OUT;
            OrderNumberTextBox.Name = "OrderNumberTextBox";
            OrderNumberTextBox.PasswordChar = '\0';
            OrderNumberTextBox.PrefixSuffixText = null;
            OrderNumberTextBox.ReadOnly = false;
            OrderNumberTextBox.RightToLeft = RightToLeft.No;
            OrderNumberTextBox.SelectedText = "";
            OrderNumberTextBox.SelectionLength = 0;
            OrderNumberTextBox.SelectionStart = 0;
            OrderNumberTextBox.ShortcutsEnabled = true;
            OrderNumberTextBox.Size = new Size(312, 48);
            OrderNumberTextBox.TabIndex = 0;
            OrderNumberTextBox.TabStop = false;
            OrderNumberTextBox.TextAlign = HorizontalAlignment.Left;
            OrderNumberTextBox.TrailingIcon = null;
            OrderNumberTextBox.UseSystemPasswordChar = false;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label1.ForeColor = Color.Navy;
            label1.Location = new Point(96, 32);
            label1.Name = "label1";
            label1.Size = new Size(202, 23);
            label1.TabIndex = 0;
            label1.Text = "Certificate Informations";
            // 
            // DATA_PAGE
            // 
            DATA_PAGE.Controls.Add(panel2);
            DATA_PAGE.Location = new Point(4, 29);
            DATA_PAGE.Name = "DATA_PAGE";
            DATA_PAGE.Padding = new Padding(3);
            DATA_PAGE.Size = new Size(1219, 721);
            DATA_PAGE.TabIndex = 1;
            DATA_PAGE.Text = "tabPage2";
            DATA_PAGE.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            panel2.BackColor = Color.White;
            panel2.Controls.Add(groupBox10);
            panel2.Controls.Add(groupBox5);
            panel2.Controls.Add(BackBox2);
            panel2.Dock = DockStyle.Fill;
            panel2.Location = new Point(3, 3);
            panel2.Name = "panel2";
            panel2.Size = new Size(1213, 715);
            panel2.TabIndex = 8;
            // 
            // groupBox10
            // 
            groupBox10.Controls.Add(CreateCertificate_Button);
            groupBox10.Controls.Add(groupBox11);
            groupBox10.Location = new Point(779, 55);
            groupBox10.Name = "groupBox10";
            groupBox10.Size = new Size(406, 665);
            groupBox10.TabIndex = 11;
            groupBox10.TabStop = false;
            // 
            // CreateCertificate_Button
            // 
            CreateCertificate_Button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            CreateCertificate_Button.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            CreateCertificate_Button.Depth = 0;
            CreateCertificate_Button.HighEmphasis = true;
            CreateCertificate_Button.Icon = null;
            CreateCertificate_Button.Location = new Point(126, 529);
            CreateCertificate_Button.Margin = new Padding(4, 6, 4, 6);
            CreateCertificate_Button.MouseState = MaterialSkin.MouseState.HOVER;
            CreateCertificate_Button.Name = "CreateCertificate_Button";
            CreateCertificate_Button.NoAccentTextColor = Color.Empty;
            CreateCertificate_Button.Size = new Size(169, 36);
            CreateCertificate_Button.TabIndex = 1;
            CreateCertificate_Button.Text = "CREATE CERTİFİCATE";
            CreateCertificate_Button.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            CreateCertificate_Button.UseAccentColor = false;
            CreateCertificate_Button.UseVisualStyleBackColor = true;
            CreateCertificate_Button.Click += CreateCertificate_Button_Click;
            // 
            // groupBox11
            // 
            groupBox11.Controls.Add(ReceiveData_Button);
            groupBox11.Controls.Add(SelectExcel_Button);
            groupBox11.Controls.Add(ExcelPage_ComboBox);
            groupBox11.Controls.Add(ExcelFileName_TextBox);
            groupBox11.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            groupBox11.ForeColor = Color.Navy;
            groupBox11.Location = new Point(26, 129);
            groupBox11.Name = "groupBox11";
            groupBox11.Size = new Size(354, 334);
            groupBox11.TabIndex = 0;
            groupBox11.TabStop = false;
            groupBox11.Text = "Excel File Selection";
            // 
            // ReceiveData_Button
            // 
            ReceiveData_Button.AutoSize = false;
            ReceiveData_Button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            ReceiveData_Button.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            ReceiveData_Button.Depth = 0;
            ReceiveData_Button.HighEmphasis = true;
            ReceiveData_Button.Icon = null;
            ReceiveData_Button.Location = new Point(197, 235);
            ReceiveData_Button.Margin = new Padding(4, 6, 4, 6);
            ReceiveData_Button.MouseState = MaterialSkin.MouseState.HOVER;
            ReceiveData_Button.Name = "ReceiveData_Button";
            ReceiveData_Button.NoAccentTextColor = Color.Empty;
            ReceiveData_Button.Size = new Size(122, 42);
            ReceiveData_Button.TabIndex = 3;
            ReceiveData_Button.Text = "RECEİVE DATA";
            ReceiveData_Button.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            ReceiveData_Button.UseAccentColor = false;
            ReceiveData_Button.UseVisualStyleBackColor = true;
            ReceiveData_Button.Click += ReceiveData_Button_Click;
            // 
            // SelectExcel_Button
            // 
            SelectExcel_Button.AutoSize = false;
            SelectExcel_Button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            SelectExcel_Button.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            SelectExcel_Button.Depth = 0;
            SelectExcel_Button.HighEmphasis = true;
            SelectExcel_Button.Icon = null;
            SelectExcel_Button.Location = new Point(46, 235);
            SelectExcel_Button.Margin = new Padding(4, 6, 4, 6);
            SelectExcel_Button.MouseState = MaterialSkin.MouseState.HOVER;
            SelectExcel_Button.Name = "SelectExcel_Button";
            SelectExcel_Button.NoAccentTextColor = Color.Empty;
            SelectExcel_Button.Size = new Size(122, 42);
            SelectExcel_Button.TabIndex = 2;
            SelectExcel_Button.Text = "SELECT FİLE";
            SelectExcel_Button.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            SelectExcel_Button.UseAccentColor = false;
            SelectExcel_Button.UseVisualStyleBackColor = true;
            SelectExcel_Button.Click += SelectExcel_Button_Click;
            // 
            // ExcelPage_ComboBox
            // 
            ExcelPage_ComboBox.AutoResize = false;
            ExcelPage_ComboBox.BackColor = Color.FromArgb(255, 255, 255);
            ExcelPage_ComboBox.Depth = 0;
            ExcelPage_ComboBox.DrawMode = DrawMode.OwnerDrawVariable;
            ExcelPage_ComboBox.DropDownHeight = 174;
            ExcelPage_ComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            ExcelPage_ComboBox.DropDownWidth = 121;
            ExcelPage_ComboBox.Font = new Font("Microsoft Sans Serif", 14F, FontStyle.Bold, GraphicsUnit.Pixel);
            ExcelPage_ComboBox.ForeColor = Color.FromArgb(222, 0, 0, 0);
            ExcelPage_ComboBox.FormattingEnabled = true;
            ExcelPage_ComboBox.Hint = "Please Select Excel Page";
            ExcelPage_ComboBox.IntegralHeight = false;
            ExcelPage_ComboBox.ItemHeight = 43;
            ExcelPage_ComboBox.Location = new Point(21, 131);
            ExcelPage_ComboBox.MaxDropDownItems = 4;
            ExcelPage_ComboBox.MouseState = MaterialSkin.MouseState.OUT;
            ExcelPage_ComboBox.Name = "ExcelPage_ComboBox";
            ExcelPage_ComboBox.Size = new Size(312, 49);
            ExcelPage_ComboBox.StartIndex = 0;
            ExcelPage_ComboBox.TabIndex = 1;
            ExcelPage_ComboBox.SelectedIndexChanged += ExcelPage_ComboBox_SelectedIndexChanged;
            // 
            // ExcelFileName_TextBox
            // 
            ExcelFileName_TextBox.AnimateReadOnly = false;
            ExcelFileName_TextBox.BackgroundImageLayout = ImageLayout.None;
            ExcelFileName_TextBox.CharacterCasing = CharacterCasing.Normal;
            ExcelFileName_TextBox.Depth = 0;
            ExcelFileName_TextBox.Font = new Font("Microsoft Sans Serif", 16F, FontStyle.Regular, GraphicsUnit.Pixel);
            ExcelFileName_TextBox.HideSelection = true;
            ExcelFileName_TextBox.Hint = "Excel File Name";
            ExcelFileName_TextBox.LeadingIcon = null;
            ExcelFileName_TextBox.Location = new Point(21, 46);
            ExcelFileName_TextBox.MaxLength = 32767;
            ExcelFileName_TextBox.MouseState = MaterialSkin.MouseState.OUT;
            ExcelFileName_TextBox.Name = "ExcelFileName_TextBox";
            ExcelFileName_TextBox.PasswordChar = '\0';
            ExcelFileName_TextBox.PrefixSuffixText = null;
            ExcelFileName_TextBox.ReadOnly = false;
            ExcelFileName_TextBox.RightToLeft = RightToLeft.No;
            ExcelFileName_TextBox.SelectedText = "";
            ExcelFileName_TextBox.SelectionLength = 0;
            ExcelFileName_TextBox.SelectionStart = 0;
            ExcelFileName_TextBox.ShortcutsEnabled = true;
            ExcelFileName_TextBox.Size = new Size(312, 48);
            ExcelFileName_TextBox.TabIndex = 0;
            ExcelFileName_TextBox.TabStop = false;
            ExcelFileName_TextBox.TextAlign = HorizontalAlignment.Left;
            ExcelFileName_TextBox.TrailingIcon = null;
            ExcelFileName_TextBox.UseSystemPasswordChar = false;
            // 
            // groupBox5
            // 
            groupBox5.Controls.Add(CheckBoxTabControl);
            groupBox5.Controls.Add(MeasurementTypes_ComboBox);
            groupBox5.Controls.Add(label3);
            groupBox5.Location = new Point(20, 55);
            groupBox5.Name = "groupBox5";
            groupBox5.Size = new Size(699, 665);
            groupBox5.TabIndex = 10;
            groupBox5.TabStop = false;
            // 
            // CheckBoxTabControl
            // 
            CheckBoxTabControl.Controls.Add(EE_Page);
            CheckBoxTabControl.Controls.Add(CalFactor_Page);
            CheckBoxTabControl.Controls.Add(CIS_Page);
            CheckBoxTabControl.Controls.Add(RFPow_Page);
            CheckBoxTabControl.Controls.Add(SParam_Page);
            CheckBoxTabControl.Controls.Add(MetCH_Page);
            CheckBoxTabControl.Controls.Add(Noise_Page);
            CheckBoxTabControl.Depth = 0;
            CheckBoxTabControl.Location = new Point(64, 129);
            CheckBoxTabControl.MouseState = MaterialSkin.MouseState.HOVER;
            CheckBoxTabControl.Multiline = true;
            CheckBoxTabControl.Name = "CheckBoxTabControl";
            CheckBoxTabControl.SelectedIndex = 0;
            CheckBoxTabControl.Size = new Size(579, 508);
            CheckBoxTabControl.TabIndex = 10;
            // 
            // EE_Page
            // 
            EE_Page.BackColor = Color.White;
            EE_Page.Controls.Add(checkBox_EE_CF);
            EE_Page.Controls.Add(checkBoxRHO);
            EE_Page.Controls.Add(checkBox_EE_RI);
            EE_Page.Controls.Add(checkBoxEE);
            EE_Page.Location = new Point(4, 29);
            EE_Page.Name = "EE_Page";
            EE_Page.Padding = new Padding(3);
            EE_Page.Size = new Size(571, 475);
            EE_Page.TabIndex = 0;
            EE_Page.Text = "Effecitve Effiency";
            // 
            // checkBox_EE_CF
            // 
            checkBox_EE_CF.AutoSize = true;
            checkBox_EE_CF.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBox_EE_CF.ForeColor = Color.Navy;
            checkBox_EE_CF.Location = new Point(66, 270);
            checkBox_EE_CF.Name = "checkBox_EE_CF";
            checkBox_EE_CF.Size = new Size(175, 24);
            checkBox_EE_CF.TabIndex = 3;
            checkBox_EE_CF.Text = "EE Calibration Factor";
            checkBox_EE_CF.UseVisualStyleBackColor = true;
            // 
            // checkBoxRHO
            // 
            checkBoxRHO.AutoSize = true;
            checkBoxRHO.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxRHO.ForeColor = Color.Navy;
            checkBoxRHO.Location = new Point(66, 206);
            checkBoxRHO.Name = "checkBoxRHO";
            checkBoxRHO.Size = new Size(84, 24);
            checkBoxRHO.TabIndex = 2;
            checkBoxRHO.Text = "Rho Lin";
            checkBoxRHO.UseVisualStyleBackColor = true;
            // 
            // checkBox_EE_RI
            // 
            checkBox_EE_RI.AutoSize = true;
            checkBox_EE_RI.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBox_EE_RI.ForeColor = Color.Navy;
            checkBox_EE_RI.Location = new Point(66, 142);
            checkBox_EE_RI.Name = "checkBox_EE_RI";
            checkBox_EE_RI.Size = new Size(158, 24);
            checkBox_EE_RI.TabIndex = 1;
            checkBox_EE_RI.Text = "Reel and Imaginer";
            checkBox_EE_RI.UseVisualStyleBackColor = true;
            // 
            // checkBoxEE
            // 
            checkBoxEE.AutoSize = true;
            checkBoxEE.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxEE.ForeColor = Color.Navy;
            checkBoxEE.Location = new Point(66, 78);
            checkBoxEE.Name = "checkBoxEE";
            checkBoxEE.Size = new Size(152, 24);
            checkBoxEE.TabIndex = 0;
            checkBoxEE.Text = "Effective Effiency";
            checkBoxEE.UseVisualStyleBackColor = true;
            // 
            // CalFactor_Page
            // 
            CalFactor_Page.Controls.Add(CF_checkBox_RIRC);
            CalFactor_Page.Controls.Add(CheckBox_CF);
            CalFactor_Page.Location = new Point(4, 29);
            CalFactor_Page.Name = "CalFactor_Page";
            CalFactor_Page.Padding = new Padding(3);
            CalFactor_Page.Size = new Size(571, 475);
            CalFactor_Page.TabIndex = 1;
            CalFactor_Page.Text = "Cal Factor";
            CalFactor_Page.UseVisualStyleBackColor = true;
            // 
            // CF_checkBox_RIRC
            // 
            CF_checkBox_RIRC.AutoSize = true;
            CF_checkBox_RIRC.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            CF_checkBox_RIRC.ForeColor = Color.Navy;
            CF_checkBox_RIRC.Location = new Point(66, 142);
            CF_checkBox_RIRC.Name = "CF_checkBox_RIRC";
            CF_checkBox_RIRC.Size = new Size(317, 24);
            CF_checkBox_RIRC.TabIndex = 5;
            CF_checkBox_RIRC.Text = "Reel , Imaginer and Reflection Coffiecent";
            CF_checkBox_RIRC.UseVisualStyleBackColor = true;
            // 
            // CheckBox_CF
            // 
            CheckBox_CF.AutoSize = true;
            CheckBox_CF.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            CheckBox_CF.ForeColor = Color.Navy;
            CheckBox_CF.Location = new Point(66, 78);
            CheckBox_CF.Name = "CheckBox_CF";
            CheckBox_CF.Size = new Size(271, 24);
            CheckBox_CF.TabIndex = 4;
            CheckBox_CF.Text = "Calibration Factor and Uncertainty";
            CheckBox_CF.UseVisualStyleBackColor = true;
            // 
            // CIS_Page
            // 
            CIS_Page.Controls.Add(CIS_CheckBox);
            CIS_Page.Location = new Point(4, 29);
            CIS_Page.Name = "CIS_Page";
            CIS_Page.Padding = new Padding(3);
            CIS_Page.Size = new Size(571, 475);
            CIS_Page.TabIndex = 2;
            CIS_Page.Text = "CIS";
            CIS_Page.UseVisualStyleBackColor = true;
            // 
            // CIS_CheckBox
            // 
            CIS_CheckBox.AutoSize = true;
            CIS_CheckBox.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            CIS_CheckBox.ForeColor = Color.Navy;
            CIS_CheckBox.Location = new Point(66, 78);
            CIS_CheckBox.Name = "CIS_CheckBox";
            CIS_CheckBox.Size = new Size(199, 24);
            CIS_CheckBox.TabIndex = 5;
            CIS_CheckBox.Text = "Z-Position , ICOD , OCID";
            CIS_CheckBox.UseVisualStyleBackColor = true;
            // 
            // RFPow_Page
            // 
            RFPow_Page.Location = new Point(4, 29);
            RFPow_Page.Name = "RFPow_Page";
            RFPow_Page.Padding = new Padding(3);
            RFPow_Page.Size = new Size(571, 475);
            RFPow_Page.TabIndex = 3;
            RFPow_Page.Text = "RF Power";
            RFPow_Page.UseVisualStyleBackColor = true;
            // 
            // SParam_Page
            // 
            SParam_Page.Controls.Add(groupBox9);
            SParam_Page.Controls.Add(groupBox8);
            SParam_Page.Controls.Add(groupBox7);
            SParam_Page.Controls.Add(groupBox6);
            SParam_Page.Location = new Point(4, 29);
            SParam_Page.Name = "SParam_Page";
            SParam_Page.Padding = new Padding(3);
            SParam_Page.Size = new Size(571, 475);
            SParam_Page.TabIndex = 4;
            SParam_Page.Text = "S-Parameter";
            SParam_Page.UseVisualStyleBackColor = true;
            // 
            // groupBox9
            // 
            groupBox9.Controls.Add(checkBoxS22SWR);
            groupBox9.Controls.Add(checkBoxS22Log);
            groupBox9.Controls.Add(checkBoxS22Lin);
            groupBox9.Controls.Add(checkBoxS22Reel);
            groupBox9.Location = new Point(36, 350);
            groupBox9.Name = "groupBox9";
            groupBox9.Size = new Size(499, 92);
            groupBox9.TabIndex = 3;
            groupBox9.TabStop = false;
            groupBox9.Text = "S22";
            // 
            // checkBoxS22SWR
            // 
            checkBoxS22SWR.AutoSize = true;
            checkBoxS22SWR.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS22SWR.ForeColor = Color.Navy;
            checkBoxS22SWR.Location = new Point(284, 62);
            checkBoxS22SWR.Name = "checkBoxS22SWR";
            checkBoxS22SWR.Size = new Size(64, 24);
            checkBoxS22SWR.TabIndex = 9;
            checkBoxS22SWR.Text = "SWR";
            checkBoxS22SWR.UseVisualStyleBackColor = true;
            // 
            // checkBoxS22Log
            // 
            checkBoxS22Log.AutoSize = true;
            checkBoxS22Log.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS22Log.ForeColor = Color.Navy;
            checkBoxS22Log.Location = new Point(17, 62);
            checkBoxS22Log.Name = "checkBoxS22Log";
            checkBoxS22Log.Size = new Size(216, 24);
            checkBoxS22Log.TabIndex = 8;
            checkBoxS22Log.Text = "Logaritmic Mag and Phase";
            checkBoxS22Log.UseVisualStyleBackColor = true;
            // 
            // checkBoxS22Lin
            // 
            checkBoxS22Lin.AutoSize = true;
            checkBoxS22Lin.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS22Lin.ForeColor = Color.Navy;
            checkBoxS22Lin.Location = new Point(284, 26);
            checkBoxS22Lin.Name = "checkBoxS22Lin";
            checkBoxS22Lin.Size = new Size(184, 24);
            checkBoxS22Lin.TabIndex = 7;
            checkBoxS22Lin.Text = "Linear Mag and Phase";
            checkBoxS22Lin.UseVisualStyleBackColor = true;
            // 
            // checkBoxS22Reel
            // 
            checkBoxS22Reel.AutoSize = true;
            checkBoxS22Reel.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS22Reel.ForeColor = Color.Navy;
            checkBoxS22Reel.Location = new Point(17, 26);
            checkBoxS22Reel.Name = "checkBoxS22Reel";
            checkBoxS22Reel.Size = new Size(158, 24);
            checkBoxS22Reel.TabIndex = 6;
            checkBoxS22Reel.Text = "Reel and Imaginer";
            checkBoxS22Reel.UseVisualStyleBackColor = true;
            // 
            // groupBox8
            // 
            groupBox8.Controls.Add(checkBoxS21Log);
            groupBox8.Controls.Add(checkBoxS21Lin);
            groupBox8.Controls.Add(checkBoxS21Reel);
            groupBox8.Location = new Point(36, 240);
            groupBox8.Name = "groupBox8";
            groupBox8.Size = new Size(499, 92);
            groupBox8.TabIndex = 2;
            groupBox8.TabStop = false;
            groupBox8.Text = "S21";
            // 
            // checkBoxS21Log
            // 
            checkBoxS21Log.AutoSize = true;
            checkBoxS21Log.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS21Log.ForeColor = Color.Navy;
            checkBoxS21Log.Location = new Point(17, 62);
            checkBoxS21Log.Name = "checkBoxS21Log";
            checkBoxS21Log.Size = new Size(216, 24);
            checkBoxS21Log.TabIndex = 8;
            checkBoxS21Log.Text = "Logaritmic Mag and Phase";
            checkBoxS21Log.UseVisualStyleBackColor = true;
            // 
            // checkBoxS21Lin
            // 
            checkBoxS21Lin.AutoSize = true;
            checkBoxS21Lin.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS21Lin.ForeColor = Color.Navy;
            checkBoxS21Lin.Location = new Point(284, 26);
            checkBoxS21Lin.Name = "checkBoxS21Lin";
            checkBoxS21Lin.Size = new Size(184, 24);
            checkBoxS21Lin.TabIndex = 7;
            checkBoxS21Lin.Text = "Linear Mag and Phase";
            checkBoxS21Lin.UseVisualStyleBackColor = true;
            // 
            // checkBoxS21Reel
            // 
            checkBoxS21Reel.AutoSize = true;
            checkBoxS21Reel.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS21Reel.ForeColor = Color.Navy;
            checkBoxS21Reel.Location = new Point(17, 26);
            checkBoxS21Reel.Name = "checkBoxS21Reel";
            checkBoxS21Reel.Size = new Size(158, 24);
            checkBoxS21Reel.TabIndex = 6;
            checkBoxS21Reel.Text = "Reel and Imaginer";
            checkBoxS21Reel.UseVisualStyleBackColor = true;
            // 
            // groupBox7
            // 
            groupBox7.Controls.Add(checkBoxS12Log);
            groupBox7.Controls.Add(checkBoxS12Lin);
            groupBox7.Controls.Add(checkBoxS12Reel);
            groupBox7.Location = new Point(36, 130);
            groupBox7.Name = "groupBox7";
            groupBox7.Size = new Size(499, 92);
            groupBox7.TabIndex = 1;
            groupBox7.TabStop = false;
            groupBox7.Text = "S12";
            // 
            // checkBoxS12Log
            // 
            checkBoxS12Log.AutoSize = true;
            checkBoxS12Log.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS12Log.ForeColor = Color.Navy;
            checkBoxS12Log.Location = new Point(17, 62);
            checkBoxS12Log.Name = "checkBoxS12Log";
            checkBoxS12Log.Size = new Size(216, 24);
            checkBoxS12Log.TabIndex = 8;
            checkBoxS12Log.Text = "Logaritmic Mag and Phase";
            checkBoxS12Log.UseVisualStyleBackColor = true;
            // 
            // checkBoxS12Lin
            // 
            checkBoxS12Lin.AutoSize = true;
            checkBoxS12Lin.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS12Lin.ForeColor = Color.Navy;
            checkBoxS12Lin.Location = new Point(284, 26);
            checkBoxS12Lin.Name = "checkBoxS12Lin";
            checkBoxS12Lin.Size = new Size(184, 24);
            checkBoxS12Lin.TabIndex = 7;
            checkBoxS12Lin.Text = "Linear Mag and Phase";
            checkBoxS12Lin.UseVisualStyleBackColor = true;
            // 
            // checkBoxS12Reel
            // 
            checkBoxS12Reel.AutoSize = true;
            checkBoxS12Reel.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS12Reel.ForeColor = Color.Navy;
            checkBoxS12Reel.Location = new Point(17, 26);
            checkBoxS12Reel.Name = "checkBoxS12Reel";
            checkBoxS12Reel.Size = new Size(158, 24);
            checkBoxS12Reel.TabIndex = 6;
            checkBoxS12Reel.Text = "Reel and Imaginer";
            checkBoxS12Reel.UseVisualStyleBackColor = true;
            // 
            // groupBox6
            // 
            groupBox6.Controls.Add(checkBoxS11SWR);
            groupBox6.Controls.Add(checkBoxS11Log);
            groupBox6.Controls.Add(checkBoxS11Lin);
            groupBox6.Controls.Add(checkBoxS11Reel);
            groupBox6.Location = new Point(36, 20);
            groupBox6.Name = "groupBox6";
            groupBox6.Size = new Size(499, 92);
            groupBox6.TabIndex = 0;
            groupBox6.TabStop = false;
            groupBox6.Text = "S11";
            // 
            // checkBoxS11SWR
            // 
            checkBoxS11SWR.AutoSize = true;
            checkBoxS11SWR.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS11SWR.ForeColor = Color.Navy;
            checkBoxS11SWR.Location = new Point(284, 62);
            checkBoxS11SWR.Name = "checkBoxS11SWR";
            checkBoxS11SWR.Size = new Size(64, 24);
            checkBoxS11SWR.TabIndex = 9;
            checkBoxS11SWR.Text = "SWR";
            checkBoxS11SWR.UseVisualStyleBackColor = true;
            // 
            // checkBoxS11Log
            // 
            checkBoxS11Log.AutoSize = true;
            checkBoxS11Log.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS11Log.ForeColor = Color.Navy;
            checkBoxS11Log.Location = new Point(17, 62);
            checkBoxS11Log.Name = "checkBoxS11Log";
            checkBoxS11Log.Size = new Size(216, 24);
            checkBoxS11Log.TabIndex = 8;
            checkBoxS11Log.Text = "Logaritmic Mag and Phase";
            checkBoxS11Log.UseVisualStyleBackColor = true;
            // 
            // checkBoxS11Lin
            // 
            checkBoxS11Lin.AutoSize = true;
            checkBoxS11Lin.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS11Lin.ForeColor = Color.Navy;
            checkBoxS11Lin.Location = new Point(284, 26);
            checkBoxS11Lin.Name = "checkBoxS11Lin";
            checkBoxS11Lin.Size = new Size(184, 24);
            checkBoxS11Lin.TabIndex = 7;
            checkBoxS11Lin.Text = "Linear Mag and Phase";
            checkBoxS11Lin.UseVisualStyleBackColor = true;
            // 
            // checkBoxS11Reel
            // 
            checkBoxS11Reel.AutoSize = true;
            checkBoxS11Reel.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            checkBoxS11Reel.ForeColor = Color.Navy;
            checkBoxS11Reel.Location = new Point(17, 26);
            checkBoxS11Reel.Name = "checkBoxS11Reel";
            checkBoxS11Reel.Size = new Size(158, 24);
            checkBoxS11Reel.TabIndex = 6;
            checkBoxS11Reel.Text = "Reel and Imaginer";
            checkBoxS11Reel.UseVisualStyleBackColor = true;
            // 
            // MetCH_Page
            // 
            MetCH_Page.Location = new Point(4, 29);
            MetCH_Page.Name = "MetCH_Page";
            MetCH_Page.Padding = new Padding(3);
            MetCH_Page.Size = new Size(571, 475);
            MetCH_Page.TabIndex = 5;
            MetCH_Page.Text = "Meteral Ch.";
            MetCH_Page.UseVisualStyleBackColor = true;
            // 
            // Noise_Page
            // 
            Noise_Page.Location = new Point(4, 29);
            Noise_Page.Name = "Noise_Page";
            Noise_Page.Padding = new Padding(3);
            Noise_Page.Size = new Size(571, 475);
            Noise_Page.TabIndex = 6;
            Noise_Page.Text = "Noise";
            Noise_Page.UseVisualStyleBackColor = true;
            // 
            // MeasurementTypes_ComboBox
            // 
            MeasurementTypes_ComboBox.AutoResize = false;
            MeasurementTypes_ComboBox.BackColor = Color.FromArgb(255, 255, 255);
            MeasurementTypes_ComboBox.Depth = 0;
            MeasurementTypes_ComboBox.DrawMode = DrawMode.OwnerDrawVariable;
            MeasurementTypes_ComboBox.DropDownHeight = 174;
            MeasurementTypes_ComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            MeasurementTypes_ComboBox.DropDownWidth = 121;
            MeasurementTypes_ComboBox.Font = new Font("Microsoft Sans Serif", 14F, FontStyle.Bold, GraphicsUnit.Pixel);
            MeasurementTypes_ComboBox.ForeColor = Color.FromArgb(222, 0, 0, 0);
            MeasurementTypes_ComboBox.FormattingEnabled = true;
            MeasurementTypes_ComboBox.IntegralHeight = false;
            MeasurementTypes_ComboBox.ItemHeight = 43;
            MeasurementTypes_ComboBox.Items.AddRange(new object[] { "Effective Effiency", "Calibration Factor", "Calculable Impedans Standard", "RF Power ", "S Paramaters", "Meteral Characterization ", "Noise" });
            MeasurementTypes_ComboBox.Location = new Point(229, 60);
            MeasurementTypes_ComboBox.MaxDropDownItems = 4;
            MeasurementTypes_ComboBox.MouseState = MaterialSkin.MouseState.OUT;
            MeasurementTypes_ComboBox.Name = "MeasurementTypes_ComboBox";
            MeasurementTypes_ComboBox.Size = new Size(241, 49);
            MeasurementTypes_ComboBox.StartIndex = 0;
            MeasurementTypes_ComboBox.TabIndex = 9;
            MeasurementTypes_ComboBox.SelectedIndexChanged += MeasurementTypes_ComboBox_SelectedIndexChanged;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point, 162);
            label3.ForeColor = Color.Navy;
            label3.Location = new Point(266, 23);
            label3.Name = "label3";
            label3.Size = new Size(169, 23);
            label3.TabIndex = 8;
            label3.Text = "Measurement Types";
            // 
            // BackBox2
            // 
            BackBox2.Image = (Image)resources.GetObject("BackBox2.Image");
            BackBox2.Location = new Point(5, 14);
            BackBox2.Name = "BackBox2";
            BackBox2.Size = new Size(62, 35);
            BackBox2.SizeMode = PictureBoxSizeMode.StretchImage;
            BackBox2.TabIndex = 7;
            BackBox2.TabStop = false;
            BackBox2.Click += BackBox2_Click;
            // 
            // ExcelView_Page
            // 
            ExcelView_Page.Controls.Add(panel3);
            ExcelView_Page.Location = new Point(4, 29);
            ExcelView_Page.Name = "ExcelView_Page";
            ExcelView_Page.Padding = new Padding(3);
            ExcelView_Page.Size = new Size(1219, 721);
            ExcelView_Page.TabIndex = 2;
            ExcelView_Page.Text = "tabPage3";
            ExcelView_Page.UseVisualStyleBackColor = true;
            // 
            // panel3
            // 
            panel3.BackColor = Color.White;
            panel3.Controls.Add(Save_Row_Col_Button);
            panel3.Controls.Add(RowNumberTextBox);
            panel3.Controls.Add(ColumnNameTextBox);
            panel3.Controls.Add(dataGridView1);
            panel3.Controls.Add(label4);
            panel3.Controls.Add(BackBox3);
            panel3.Dock = DockStyle.Fill;
            panel3.Location = new Point(3, 3);
            panel3.Name = "panel3";
            panel3.Size = new Size(1213, 715);
            panel3.TabIndex = 0;
            // 
            // Save_Row_Col_Button
            // 
            Save_Row_Col_Button.AutoSize = false;
            Save_Row_Col_Button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Save_Row_Col_Button.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            Save_Row_Col_Button.Depth = 0;
            Save_Row_Col_Button.HighEmphasis = true;
            Save_Row_Col_Button.Icon = null;
            Save_Row_Col_Button.Location = new Point(941, 661);
            Save_Row_Col_Button.Margin = new Padding(4, 6, 4, 6);
            Save_Row_Col_Button.MouseState = MaterialSkin.MouseState.HOVER;
            Save_Row_Col_Button.Name = "Save_Row_Col_Button";
            Save_Row_Col_Button.NoAccentTextColor = Color.Empty;
            Save_Row_Col_Button.Size = new Size(198, 43);
            Save_Row_Col_Button.TabIndex = 13;
            Save_Row_Col_Button.Text = "KAYDET VE ÇIK";
            Save_Row_Col_Button.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            Save_Row_Col_Button.UseAccentColor = false;
            Save_Row_Col_Button.UseVisualStyleBackColor = true;
            Save_Row_Col_Button.Click += Save_Row_Col_Button_Click;
            // 
            // RowNumberTextBox
            // 
            RowNumberTextBox.AnimateReadOnly = false;
            RowNumberTextBox.BackgroundImageLayout = ImageLayout.None;
            RowNumberTextBox.CharacterCasing = CharacterCasing.Normal;
            RowNumberTextBox.Depth = 0;
            RowNumberTextBox.Font = new Font("Microsoft Sans Serif", 16F, FontStyle.Regular, GraphicsUnit.Pixel);
            RowNumberTextBox.HideSelection = true;
            RowNumberTextBox.Hint = "Lütfen satır numarasını giriniz";
            RowNumberTextBox.LeadingIcon = null;
            RowNumberTextBox.Location = new Point(491, 661);
            RowNumberTextBox.MaxLength = 32767;
            RowNumberTextBox.MouseState = MaterialSkin.MouseState.OUT;
            RowNumberTextBox.Name = "RowNumberTextBox";
            RowNumberTextBox.PasswordChar = '\0';
            RowNumberTextBox.PrefixSuffixText = null;
            RowNumberTextBox.ReadOnly = false;
            RowNumberTextBox.RightToLeft = RightToLeft.No;
            RowNumberTextBox.SelectedText = "";
            RowNumberTextBox.SelectionLength = 0;
            RowNumberTextBox.SelectionStart = 0;
            RowNumberTextBox.ShortcutsEnabled = true;
            RowNumberTextBox.Size = new Size(312, 48);
            RowNumberTextBox.TabIndex = 12;
            RowNumberTextBox.TabStop = false;
            RowNumberTextBox.TextAlign = HorizontalAlignment.Left;
            RowNumberTextBox.TrailingIcon = null;
            RowNumberTextBox.UseSystemPasswordChar = false;
            // 
            // ColumnNameTextBox
            // 
            ColumnNameTextBox.AnimateReadOnly = false;
            ColumnNameTextBox.BackgroundImageLayout = ImageLayout.None;
            ColumnNameTextBox.CharacterCasing = CharacterCasing.Normal;
            ColumnNameTextBox.Depth = 0;
            ColumnNameTextBox.Font = new Font("Microsoft Sans Serif", 16F, FontStyle.Regular, GraphicsUnit.Pixel);
            ColumnNameTextBox.HideSelection = true;
            ColumnNameTextBox.Hint = "Lütfen Sütun adını giriniz";
            ColumnNameTextBox.LeadingIcon = null;
            ColumnNameTextBox.Location = new Point(75, 661);
            ColumnNameTextBox.MaxLength = 32767;
            ColumnNameTextBox.MouseState = MaterialSkin.MouseState.OUT;
            ColumnNameTextBox.Name = "ColumnNameTextBox";
            ColumnNameTextBox.PasswordChar = '\0';
            ColumnNameTextBox.PrefixSuffixText = null;
            ColumnNameTextBox.ReadOnly = false;
            ColumnNameTextBox.RightToLeft = RightToLeft.No;
            ColumnNameTextBox.SelectedText = "";
            ColumnNameTextBox.SelectionLength = 0;
            ColumnNameTextBox.SelectionStart = 0;
            ColumnNameTextBox.ShortcutsEnabled = true;
            ColumnNameTextBox.Size = new Size(312, 48);
            ColumnNameTextBox.TabIndex = 11;
            ColumnNameTextBox.TabStop = false;
            ColumnNameTextBox.TextAlign = HorizontalAlignment.Left;
            ColumnNameTextBox.TrailingIcon = null;
            ColumnNameTextBox.UseSystemPasswordChar = false;
            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(50, 113);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.ReadOnly = true;
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.Size = new Size(1113, 505);
            dataGridView1.TabIndex = 10;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Segoe UI", 13.8F, FontStyle.Bold, GraphicsUnit.Point, 162);
            label4.ForeColor = Color.Navy;
            label4.Location = new Point(270, 53);
            label4.Name = "label4";
            label4.Size = new Size(673, 31);
            label4.TabIndex = 9;
            label4.Text = "Lütfen satır ve sütun değerlerini aşağıdaki kutucuklara giriniz";
            // 
            // BackBox3
            // 
            BackBox3.Image = (Image)resources.GetObject("BackBox3.Image");
            BackBox3.Location = new Point(5, 14);
            BackBox3.Name = "BackBox3";
            BackBox3.Size = new Size(62, 35);
            BackBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            BackBox3.TabIndex = 8;
            BackBox3.TabStop = false;
            BackBox3.Click += BackBox3_Click;
            // 
            // progressBar
            // 
            progressBar.Location = new Point(438, 20);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(125, 29);
            progressBar.TabIndex = 0;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point, 162);
            label6.ForeColor = Color.Navy;
            label6.Location = new Point(19, 20);
            label6.Name = "label6";
            label6.Size = new Size(344, 23);
            label6.TabIndex = 2;
            label6.Text = "TÜBİTAK ULUSAL METROLOJİ ENSTİTÜSÜ";
            // 
            // groupBox12
            // 
            groupBox12.BackColor = Color.White;
            groupBox12.Controls.Add(LabelProgress);
            groupBox12.Controls.Add(label6);
            groupBox12.Controls.Add(progressBar);
            groupBox12.Location = new Point(0, 706);
            groupBox12.Name = "groupBox12";
            groupBox12.Size = new Size(1223, 60);
            groupBox12.TabIndex = 9;
            groupBox12.TabStop = false;
            // 
            // LabelProgress
            // 
            LabelProgress.AutoSize = true;
            LabelProgress.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point, 162);
            LabelProgress.Location = new Point(599, 26);
            LabelProgress.Name = "LabelProgress";
            LabelProgress.Size = new Size(121, 23);
            LabelProgress.TabIndex = 3;
            LabelProgress.Text = "LabelProgress";
            LabelProgress.Visible = false;
            // 
            // CertificateForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            ClientSize = new Size(1223, 778);
            Controls.Add(groupBox12);
            Controls.Add(CertificateTabControl);
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 162);
            Name = "CertificateForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "API OPERATİONS";
            TransparencyKey = SystemColors.Control;
            FormClosed += CertificateForm_FormClosed;
            ((System.ComponentModel.ISupportInitialize)BackBox1).EndInit();
            CertificateTabControl.ResumeLayout(false);
            API_PAGE.ResumeLayout(false);
            panel1.ResumeLayout(false);
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox3.ResumeLayout(false);
            groupBox2.ResumeLayout(false);
            DATA_PAGE.ResumeLayout(false);
            panel2.ResumeLayout(false);
            groupBox10.ResumeLayout(false);
            groupBox10.PerformLayout();
            groupBox11.ResumeLayout(false);
            groupBox5.ResumeLayout(false);
            groupBox5.PerformLayout();
            CheckBoxTabControl.ResumeLayout(false);
            EE_Page.ResumeLayout(false);
            EE_Page.PerformLayout();
            CalFactor_Page.ResumeLayout(false);
            CalFactor_Page.PerformLayout();
            CIS_Page.ResumeLayout(false);
            CIS_Page.PerformLayout();
            SParam_Page.ResumeLayout(false);
            groupBox9.ResumeLayout(false);
            groupBox9.PerformLayout();
            groupBox8.ResumeLayout(false);
            groupBox8.PerformLayout();
            groupBox7.ResumeLayout(false);
            groupBox7.PerformLayout();
            groupBox6.ResumeLayout(false);
            groupBox6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)BackBox2).EndInit();
            ExcelView_Page.ResumeLayout(false);
            panel3.ResumeLayout(false);
            panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)BackBox3).EndInit();
            groupBox12.ResumeLayout(false);
            groupBox12.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private PictureBox BackBox1;
        private TabControl CertificateTabControl;
        private TabPage API_PAGE;
        private Panel panel1;
        private TabPage DATA_PAGE;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private Label label1;
        private GroupBox groupBox3;
        private MaterialSkin.Controls.MaterialComboBox LaboratoryComboBox;
        private MaterialSkin.Controls.MaterialTextBox2 OrderNumberTextBox;
        private MaterialSkin.Controls.MaterialTextBox2 CalCodeTextBox;
        private MaterialSkin.Controls.MaterialTextBox2 SerialNumberTextBox;
        private MaterialSkin.Controls.MaterialTextBox2 ModelNameTextBox;
        private MaterialSkin.Controls.MaterialTextBox2 DeviceNameTextBox;
        private MaterialSkin.Controls.MaterialButton SelectDeviceButton;
        private GroupBox groupBox4;
        private Label label2;
        private MaterialSkin.Controls.MaterialMultiLineTextBox2 MeasurementsTextBox;
        private MaterialSkin.Controls.MaterialMultiLineTextBox2 CalibrationDescTextBox;
        private MaterialSkin.Controls.MaterialMultiLineTextBox2 MethodTextBox;
        private MaterialSkin.Controls.MaterialMultiLineTextBox2 DeviceTextBox;
        private MaterialSkin.Controls.MaterialButton NextButton;
        private PictureBox BackBox2;
        private Panel panel2;
        private GroupBox groupBox5;
        private MaterialSkin.Controls.MaterialTabControl CheckBoxTabControl;
        private TabPage EE_Page;
        private TabPage CalFactor_Page;
        private MaterialSkin.Controls.MaterialComboBox MeasurementTypes_ComboBox;
        private Label label3;
        private TabPage ExcelView_Page;
        private TabPage CIS_Page;
        private TabPage RFPow_Page;
        private TabPage SParam_Page;
        private TabPage MetCH_Page;
        private TabPage Noise_Page;
        private CheckBox checkBox_EE_CF;
        private CheckBox checkBoxRHO;
        private CheckBox checkBox_EE_RI;
        private CheckBox checkBoxEE;
        private CheckBox CF_checkBox_RIRC;
        private CheckBox CheckBox_CF;
        private CheckBox CIS_CheckBox;
        private GroupBox groupBox6;
        private GroupBox groupBox8;
        private GroupBox groupBox7;
        private GroupBox groupBox9;
        private GroupBox groupBox10;
        private GroupBox groupBox11;
        private MaterialSkin.Controls.MaterialButton CreateCertificate_Button;
        private MaterialSkin.Controls.MaterialButton ReceiveData_Button;
        private MaterialSkin.Controls.MaterialButton SelectExcel_Button;
        private MaterialSkin.Controls.MaterialComboBox ExcelPage_ComboBox;
        private Panel panel3;
        private PictureBox BackBox3;
        private DataGridView dataGridView1;
        private Label label4;
        private MaterialSkin.Controls.MaterialButton Save_Row_Col_Button;
        private MaterialSkin.Controls.MaterialTextBox2 RowNumberTextBox;
        private MaterialSkin.Controls.MaterialTextBox2 ColumnNameTextBox;
        private Label label6;
        private GroupBox groupBox12;
        public MaterialSkin.Controls.MaterialTextBox2 ExcelFileName_TextBox;
        public CheckBox checkBoxS21Log;
        public CheckBox checkBoxS21Lin;
        public CheckBox checkBoxS21Reel;
        public CheckBox checkBoxS12Log;
        public CheckBox checkBoxS12Lin;
        public CheckBox checkBoxS12Reel;
        public CheckBox checkBoxS11SWR;
        public CheckBox checkBoxS11Log;
        public CheckBox checkBoxS11Lin;
        public CheckBox checkBoxS11Reel;
        public CheckBox checkBoxS22SWR;
        public CheckBox checkBoxS22Log;
        public CheckBox checkBoxS22Lin;
        public CheckBox checkBoxS22Reel;
        public ProgressBar progressBar;
        public Label LabelProgress;
    }
}