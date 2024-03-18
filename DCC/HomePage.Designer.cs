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
            materialButton1 = new MaterialSkin.Controls.MaterialButton();
            materialTextBox22 = new MaterialSkin.Controls.MaterialTextBox2();
            materialButton3 = new MaterialSkin.Controls.MaterialButton();
            pictureBox3 = new PictureBox();
            HomeTab = new TabPage();
            panel1 = new Panel();
            XMLtoWordPageButton = new MaterialSkin.Controls.MaterialButton();
            CertificatePageShowButton = new MaterialSkin.Controls.MaterialButton();
            pictureBox1 = new PictureBox();
            tabControl1 = new TabControl();
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
            XMLtoWordPage.Size = new Size(697, 707);
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
            labelProgress.Location = new Point(372, 19);
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
            panel2.Size = new Size(691, 701);
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
            groupBox1.Controls.Add(materialButton1);
            groupBox1.Controls.Add(materialTextBox22);
            groupBox1.Controls.Add(materialButton3);
            groupBox1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 162);
            groupBox1.ForeColor = Color.Navy;
            groupBox1.Location = new Point(162, 325);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(353, 308);
            groupBox1.TabIndex = 3;
            groupBox1.TabStop = false;
            groupBox1.Text = "XML TO WORD";
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
            // materialButton1
            // 
            materialButton1.AutoSize = false;
            materialButton1.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            materialButton1.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            materialButton1.Depth = 0;
            materialButton1.HighEmphasis = true;
            materialButton1.Icon = null;
            materialButton1.Location = new Point(99, 226);
            materialButton1.Margin = new Padding(4, 6, 4, 6);
            materialButton1.MouseState = MaterialSkin.MouseState.HOVER;
            materialButton1.Name = "materialButton1";
            materialButton1.NoAccentTextColor = Color.Empty;
            materialButton1.Size = new Size(158, 36);
            materialButton1.TabIndex = 3;
            materialButton1.Text = "CREATE WORD FİLE";
            materialButton1.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            materialButton1.UseAccentColor = false;
            materialButton1.UseVisualStyleBackColor = true;
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
            // materialButton3
            // 
            materialButton3.AutoSize = false;
            materialButton3.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            materialButton3.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            materialButton3.Depth = 0;
            materialButton3.HighEmphasis = true;
            materialButton3.Icon = null;
            materialButton3.Location = new Point(99, 159);
            materialButton3.Margin = new Padding(4, 6, 4, 6);
            materialButton3.MouseState = MaterialSkin.MouseState.HOVER;
            materialButton3.Name = "materialButton3";
            materialButton3.NoAccentTextColor = Color.Empty;
            materialButton3.Size = new Size(158, 36);
            materialButton3.TabIndex = 2;
            materialButton3.Text = "SELECT XML FİLE";
            materialButton3.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            materialButton3.UseAccentColor = false;
            materialButton3.UseVisualStyleBackColor = true;
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
            HomeTab.Size = new Size(697, 707);
            HomeTab.TabIndex = 0;
            HomeTab.Text = "tabPage1";
            // 
            // panel1
            // 
            panel1.BackColor = Color.White;
            panel1.Controls.Add(XMLtoWordPageButton);
            panel1.Controls.Add(CertificatePageShowButton);
            panel1.Controls.Add(pictureBox1);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(3, 3);
            panel1.Name = "panel1";
            panel1.Size = new Size(691, 701);
            panel1.TabIndex = 1;
            // 
            // XMLtoWordPageButton
            // 
            XMLtoWordPageButton.AutoSize = false;
            XMLtoWordPageButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            XMLtoWordPageButton.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            XMLtoWordPageButton.Depth = 0;
            XMLtoWordPageButton.HighEmphasis = true;
            XMLtoWordPageButton.Icon = null;
            XMLtoWordPageButton.Location = new Point(375, 438);
            XMLtoWordPageButton.Margin = new Padding(4, 6, 4, 6);
            XMLtoWordPageButton.MouseState = MaterialSkin.MouseState.HOVER;
            XMLtoWordPageButton.Name = "XMLtoWordPageButton";
            XMLtoWordPageButton.NoAccentTextColor = Color.Empty;
            XMLtoWordPageButton.Size = new Size(173, 66);
            XMLtoWordPageButton.TabIndex = 2;
            XMLtoWordPageButton.Text = "MACHİNE READABLE TO HUMAN READABLE";
            XMLtoWordPageButton.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            XMLtoWordPageButton.UseAccentColor = false;
            XMLtoWordPageButton.UseVisualStyleBackColor = true;
            XMLtoWordPageButton.Click += XMLtoWordPageButton_Click;
            // 
            // CertificatePageShowButton
            // 
            CertificatePageShowButton.AutoSize = false;
            CertificatePageShowButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            CertificatePageShowButton.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            CertificatePageShowButton.Depth = 0;
            CertificatePageShowButton.HighEmphasis = true;
            CertificatePageShowButton.Icon = null;
            CertificatePageShowButton.Location = new Point(109, 438);
            CertificatePageShowButton.Margin = new Padding(4, 6, 4, 6);
            CertificatePageShowButton.MouseState = MaterialSkin.MouseState.HOVER;
            CertificatePageShowButton.Name = "CertificatePageShowButton";
            CertificatePageShowButton.NoAccentTextColor = Color.Empty;
            CertificatePageShowButton.Size = new Size(173, 66);
            CertificatePageShowButton.TabIndex = 1;
            CertificatePageShowButton.Text = "CREATE CERTİFİCATE";
            CertificatePageShowButton.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            CertificatePageShowButton.UseAccentColor = false;
            CertificatePageShowButton.UseVisualStyleBackColor = true;
            CertificatePageShowButton.Click += CertificatePageShowButton_Click;
            // 
            // pictureBox1
            // 
            pictureBox1.Image = (Image)resources.GetObject("pictureBox1.Image");
            pictureBox1.Location = new Point(209, 16);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(259, 280);
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.TabIndex = 0;
            pictureBox1.TabStop = false;
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(HomeTab);
            tabControl1.Controls.Add(XMLtoWordPage);
            tabControl1.Location = new Point(0, -36);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(705, 740);
            tabControl1.TabIndex = 1;
            // 
            // HomePage
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            ClientSize = new Size(699, 706);
            Controls.Add(tabControl1);
            Name = "HomePage";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "DCC CREATER ";
            FormClosed += HomePage_FormClosed;
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
        private MaterialSkin.Controls.MaterialButton materialButton1;
        private MaterialSkin.Controls.MaterialTextBox2 materialTextBox22;
        private MaterialSkin.Controls.MaterialButton materialButton3;
        private PictureBox pictureBox3;
        private TabPage HomeTab;
        private Panel panel1;
        private MaterialSkin.Controls.MaterialButton XMLtoWordPageButton;
        private MaterialSkin.Controls.MaterialButton CertificatePageShowButton;
        private PictureBox pictureBox1;
        private TabControl tabControl1;
    }
}
