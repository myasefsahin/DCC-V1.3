namespace DCC
{
    partial class HomePage
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HomePage));
            XMLtoWordPage = new TabPage();
            groupBox2 = new GroupBox();
            progressBar = new ProgressBar();
            labelProgress = new Label();
            panel2 = new Panel();
            pictureBox2 = new PictureBox();
            groupBox1 = new GroupBox();
            checkBox1 = new CheckBox();
            createWordFile = new MaterialSkin.Controls.MaterialButton();
            materialTextBox22 = new MaterialSkin.Controls.MaterialTextBox2();
            selectXmlFile = new MaterialSkin.Controls.MaterialButton();
            pictureBox3 = new PictureBox();
            HomeTab = new TabPage();
            panel1 = new Panel();
            label2 = new Label();
            label1 = new Label();
            pictureBox1 = new PictureBox();
            tabControl1 = new TabControl();
            XMLtoWordPageButton = new MaterialSkin.Controls.MaterialButton();
            CertificatePageShowButton = new MaterialSkin.Controls.MaterialButton();
            XMLtoWordPage.SuspendLayout();
            groupBox2.SuspendLayout();
            panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox2).BeginInit();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox3).BeginInit();
            HomeTab.SuspendLayout();
            panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            tabControl1.SuspendLayout();
            SuspendLayout();
            // 
            // XMLtoWordPage
            // 
            XMLtoWordPage.Controls.Add(groupBox2);
            XMLtoWordPage.Controls.Add(panel2);
            XMLtoWordPage.Location = new Point(4, 29);
            XMLtoWordPage.Name = "XMLtoWordPage";
            XMLtoWordPage.Padding = new Padding(3);
            XMLtoWordPage.Size = new Size(642, 641);
            XMLtoWordPage.TabIndex = 1;
            XMLtoWordPage.Text = "tabPage2";
            XMLtoWordPage.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(progressBar);
            groupBox2.Controls.Add(labelProgress);
            groupBox2.Location = new Point(-3, 648);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(702, 53);
            groupBox2.TabIndex = 6;
            groupBox2.TabStop = false;
            // 
            // progressBar
            // 
            progressBar.BackColor = Color.White;
            progressBar.Location = new Point(526, 17);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(125, 29);
            progressBar.TabIndex = 2;
            // 
            // labelProgress
            // 
            labelProgress.AutoSize = true;
            labelProgress.Location = new Point(373, 19);
            labelProgress.Name = "labelProgress";
            labelProgress.Size = new Size(101, 20);
            labelProgress.TabIndex = 3;
            labelProgress.Text = "LabelProgress";
            labelProgress.Visible = false;
            // 
            // panel2
            // 
            panel2.Controls.Add(pictureBox2);
            panel2.Controls.Add(groupBox1);
            panel2.Controls.Add(pictureBox3);
            panel2.Dock = DockStyle.Fill;
            panel2.Location = new Point(3, 3);
            panel2.Name = "panel2";
            panel2.Size = new Size(636, 635);
            panel2.TabIndex = 6;
            // 
            // pictureBox2
            // 
            pictureBox2.Image = (Image)resources.GetObject("pictureBox2.Image");
            pictureBox2.Location = new Point(209, 16);
            pictureBox2.Name = "pictureBox2";
            pictureBox2.Size = new Size(259, 280);
            pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox2.TabIndex = 4;
            pictureBox2.TabStop = false;
            // 
            // groupBox1
            // 
            groupBox1.BackColor = Color.White;
            groupBox1.Controls.Add(checkBox1);
            groupBox1.Controls.Add(createWordFile);
            groupBox1.Controls.Add(materialTextBox22);
            groupBox1.Controls.Add(selectXmlFile);
            groupBox1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            groupBox1.ForeColor = Color.Navy;
            groupBox1.Location = new Point(162, 325);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(353, 308);
            groupBox1.TabIndex = 3;
            groupBox1.TabStop = false;
            groupBox1.Text = "MACHİNE READABLE TO HUMAN READABLE";
            // 
            // checkBox1
            // 
            checkBox1.AutoSize = true;
            checkBox1.Location = new Point(350, 300);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(106, 24);
            checkBox1.TabIndex = 4;
            checkBox1.Text = "checkBox1";
            checkBox1.UseVisualStyleBackColor = true;
            // 
            // createWordFile
            // 
            createWordFile.AutoSize = false;
            createWordFile.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            createWordFile.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            createWordFile.Depth = 0;
            createWordFile.HighEmphasis = true;
            createWordFile.Icon = null;
            createWordFile.Location = new Point(99, 227);
            createWordFile.Margin = new Padding(5);
            createWordFile.MouseState = MaterialSkin.MouseState.HOVER;
            createWordFile.Name = "createWordFile";
            createWordFile.NoAccentTextColor = Color.Empty;
            createWordFile.Size = new Size(158, 36);
            createWordFile.TabIndex = 3;
            createWordFile.Text = "CREATE HUMAN READABLE FİLE";
            createWordFile.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            createWordFile.UseAccentColor = false;
            createWordFile.UseVisualStyleBackColor = true;
            createWordFile.Click += createWordFile_Click;
            // 
            // materialTextBox22
            // 
            materialTextBox22.AnimateReadOnly = false;
            materialTextBox22.BackgroundImageLayout = ImageLayout.None;
            materialTextBox22.CharacterCasing = CharacterCasing.Normal;
            materialTextBox22.Depth = 0;
            materialTextBox22.Font = new Font("Microsoft Sans Serif", 16F, FontStyle.Regular, GraphicsUnit.Pixel);
            materialTextBox22.HideSelection = true;
            materialTextBox22.Hint = "XML file name";
            materialTextBox22.LeadingIcon = null;
            materialTextBox22.Location = new Point(35, 81);
            materialTextBox22.MaxLength = 32767;
            materialTextBox22.MouseState = MaterialSkin.MouseState.OUT;
            materialTextBox22.Name = "materialTextBox22";
            materialTextBox22.PasswordChar = '\0';
            materialTextBox22.PrefixSuffixText = null;
            materialTextBox22.ReadOnly = true;
            materialTextBox22.RightToLeft = RightToLeft.No;
            materialTextBox22.SelectedText = "";
            materialTextBox22.SelectionLength = 0;
            materialTextBox22.SelectionStart = 0;
            materialTextBox22.ShortcutsEnabled = true;
            materialTextBox22.Size = new Size(283, 48);
            materialTextBox22.TabIndex = 1;
            materialTextBox22.TabStop = false;
            materialTextBox22.TextAlign = HorizontalAlignment.Left;
            materialTextBox22.TrailingIcon = null;
            materialTextBox22.UseSystemPasswordChar = false;
            // 
            // selectXmlFile
            // 
            selectXmlFile.AutoSize = false;
            selectXmlFile.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            selectXmlFile.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            selectXmlFile.Depth = 0;
            selectXmlFile.HighEmphasis = true;
            selectXmlFile.Icon = null;
            selectXmlFile.Location = new Point(99, 159);
            selectXmlFile.Margin = new Padding(5);
            selectXmlFile.MouseState = MaterialSkin.MouseState.HOVER;
            selectXmlFile.Name = "selectXmlFile";
            selectXmlFile.NoAccentTextColor = Color.Empty;
            selectXmlFile.Size = new Size(158, 36);
            selectXmlFile.TabIndex = 2;
            selectXmlFile.Text = "SELECT XML FİLE";
            selectXmlFile.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            selectXmlFile.UseAccentColor = false;
            selectXmlFile.UseVisualStyleBackColor = true;
            selectXmlFile.Click += selectXmlFile_Click;
            // 
            // pictureBox3
            // 
            pictureBox3.Image = (Image)resources.GetObject("pictureBox3.Image");
            pictureBox3.Location = new Point(15, 16);
            pictureBox3.Name = "pictureBox3";
            pictureBox3.Size = new Size(62, 35);
            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox3.TabIndex = 5;
            pictureBox3.TabStop = false;
            pictureBox3.Click += pictureBox3_Click;
            // 
            // HomeTab
            // 
            HomeTab.BackColor = Color.White;
            HomeTab.Controls.Add(panel1);
            HomeTab.ImeMode = ImeMode.NoControl;
            HomeTab.Location = new Point(4, 29);
            HomeTab.Name = "HomeTab";
            HomeTab.Padding = new Padding(3);
            HomeTab.Size = new Size(642, 641);
            HomeTab.TabIndex = 0;
            HomeTab.Text = "tabPage1";
            // 
            // panel1
            // 
            panel1.BackColor = Color.White;
            panel1.Controls.Add(XMLtoWordPageButton);
            panel1.Controls.Add(CertificatePageShowButton);
            panel1.Controls.Add(label2);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(pictureBox1);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(3, 3);
            panel1.Margin = new Padding(0);
            panel1.Name = "panel1";
            panel1.Size = new Size(636, 635);
            panel1.TabIndex = 1;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 15F, FontStyle.Bold);
            label2.Location = new Point(339, 337);
            label2.Name = "label2";
            label2.Size = new Size(86, 35);
            label2.TabIndex = 4;
            label2.Text = "label2";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 15F, FontStyle.Italic, GraphicsUnit.Point, 162);
            label1.Location = new Point(225, 337);
            label1.Name = "label1";
            label1.Size = new Size(80, 35);
            label1.TabIndex = 3;
            label1.Text = "label1";
            // 
            // pictureBox1
            // 
            pictureBox1.Image = (Image)resources.GetObject("pictureBox1.Image");
            pictureBox1.Location = new Point(225, 35);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(200, 250);
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.TabIndex = 0;
            pictureBox1.TabStop = false;
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(HomeTab);
            tabControl1.Controls.Add(XMLtoWordPage);
            tabControl1.Location = new Point(-10, -35);
            tabControl1.Margin = new Padding(0);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(650, 674);
            tabControl1.TabIndex = 1;
            // 
            // XMLtoWordPageButton
            // 
            XMLtoWordPageButton.AutoSize = false;
            XMLtoWordPageButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            XMLtoWordPageButton.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            XMLtoWordPageButton.Depth = 0;
            XMLtoWordPageButton.HighEmphasis = true;
            XMLtoWordPageButton.Icon = null;
            XMLtoWordPageButton.Location = new Point(339, 481);
            XMLtoWordPageButton.Margin = new Padding(5);
            XMLtoWordPageButton.MouseState = MaterialSkin.MouseState.HOVER;
            XMLtoWordPageButton.Name = "XMLtoWordPageButton";
            XMLtoWordPageButton.NoAccentTextColor = Color.Empty;
            XMLtoWordPageButton.Size = new Size(246, 68);
            XMLtoWordPageButton.TabIndex = 6;
            XMLtoWordPageButton.Text = "convert certıfıcate from machıne readable to human readable";
            XMLtoWordPageButton.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            XMLtoWordPageButton.UseAccentColor = false;
            XMLtoWordPageButton.UseVisualStyleBackColor = true;
            // 
            // CertificatePageShowButton
            // 
            CertificatePageShowButton.AutoSize = false;
            CertificatePageShowButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            CertificatePageShowButton.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            CertificatePageShowButton.Depth = 0;
            CertificatePageShowButton.HighEmphasis = true;
            CertificatePageShowButton.Icon = null;
            CertificatePageShowButton.Location = new Point(59, 481);
            CertificatePageShowButton.Margin = new Padding(5);
            CertificatePageShowButton.MouseState = MaterialSkin.MouseState.HOVER;
            CertificatePageShowButton.Name = "CertificatePageShowButton";
            CertificatePageShowButton.NoAccentTextColor = Color.Empty;
            CertificatePageShowButton.Size = new Size(246, 68);
            CertificatePageShowButton.TabIndex = 5;
            CertificatePageShowButton.Text = "create certıfıcate as machıne and human readable";
            CertificatePageShowButton.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            CertificatePageShowButton.UseAccentColor = false;
            CertificatePageShowButton.UseVisualStyleBackColor = true;
            // 
            // HomePage
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            ClientSize = new Size(632, 631);
            Controls.Add(tabControl1);
            Name = "HomePage";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "DCC CREATER ";
            FormClosed += HomePage_FormClosed;
            Load += HomePage_Load;
            XMLtoWordPage.ResumeLayout(false);
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)pictureBox2).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox3).EndInit();
            HomeTab.ResumeLayout(false);
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            tabControl1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private TabPage XMLtoWordPage;
        private GroupBox groupBox2;
        private ProgressBar progressBar;
        private Label labelProgress;
        private Panel panel2;
        private PictureBox pictureBox2;
        private GroupBox groupBox1;
        private CheckBox checkBox1;
        private MaterialSkin.Controls.MaterialButton createWordFile;
        private MaterialSkin.Controls.MaterialTextBox2 materialTextBox22;
        private MaterialSkin.Controls.MaterialButton selectXmlFile;
        private PictureBox pictureBox3;
        private TabPage HomeTab;
        private Panel panel1;
        private PictureBox pictureBox1;
        private TabControl tabControl1;
        private Label label1;
        private Label label2;
        private MaterialSkin.Controls.MaterialButton XMLtoWordPageButton;
        private MaterialSkin.Controls.MaterialButton CertificatePageShowButton;
    }
}
