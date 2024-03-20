using System.Xml;

namespace DCC
{
    public partial class HomePage : Form
    {

        public string XMLDosyaYolu;
        public string XMLDosyaAdi;
        public string WordFolderPath;
        public string XMLFolderPath;
        public string CreatedFileName;
        public XmlDocument xml = new XmlDocument();
        XML_Arrays XML_Arrays = new XML_Arrays();


        public HomePage()
        {
            InitializeComponent();
        }

        private void XMLtoWordPageButton_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = true;
            this.Text = "XML TO WORD";
            tabControl1.SelectedTab = XMLtoWordPage;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.Text = "DCC CREATER";
            tabControl1.SelectedTab = HomeTab;
        }

        private void CertificatePageShowButton_Click(object sender, EventArgs e)
        {
            this.Visible = false;



            try
            {

                string xmlFilePath = "xmlData/dcc_xml_scheme_uzun.xml";
                XmlDocument x = new XmlDocument();
                x.Load(xmlFilePath);

                CreateXML ctx = new CreateXML();
                this.xml = ctx.AddAdministrativeData(x, XML_Arrays);

                // Feedback
                labelProgress.Visible = false;
                Thread.Sleep(10);
                progressBar.Value = 0;
                for (int i = 0; i < 100; i++) progressBar.Value += 1;
                labelProgress.Visible = true;
                labelProgress.ForeColor = System.Drawing.Color.Green;
                labelProgress.Text = @"Administrative data save successful";
                CertificateForm certificateForm = new CertificateForm(xml);
                this.xml = certificateForm.xml;
                certificateForm.Show();

            }
            catch (Exception err)
            {
                labelProgress.Visible = true;
                labelProgress.ForeColor = System.Drawing.Color.Red;
                labelProgress.Text = @"ERROR!: Device Information";
                MessageBox.Show(err.Message, err.StackTrace, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void HomePage_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void selectXmlFile_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();

                openFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
                openFileDialog.InitialDirectory = "C:\\Users\\" + Environment.UserName + "\\source\\repos\\VISOS2\\bin\\Debug\\xmlData\\xml";
                openFileDialog.Title = @"XML Files, Select a "".xml"" file";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    XMLDosyaYolu = openFileDialog.FileName;
                    XMLDosyaAdi = Path.GetFileName(XMLDosyaYolu);
                    string xmlMetni = File.ReadAllText(XMLDosyaYolu);
                    // CopyToXml.SelectFilledColumns(xmlMetni);
                    materialTextBox22.Text = XMLDosyaAdi;


                }
                // Feedback iþlemi
                labelProgress.Visible = true;
                labelProgress.ForeColor = System.Drawing.Color.Green;
                labelProgress.Text = @"XML file selection successful";
            }
            catch (Exception err)
            {
                labelProgress.Visible = true;
                labelProgress.ForeColor = System.Drawing.Color.Red;
                labelProgress.Text = @"ERROR!: Selection of XML";
                MessageBox.Show(err.Message, err.StackTrace, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void createWordFile_Click(object sender, EventArgs e)
        {
            

            labelProgress.Visible = false;
            progressBar.Value = 0;

            try
            {
                XmlToWord xmlToWord = new XmlToWord();
                xmlToWord.Try(XMLDosyaYolu);
             
            }
            catch (Exception err)
            {
                labelProgress.Visible = true;
                labelProgress.ForeColor = System.Drawing.Color.Red;
                labelProgress.Text = @"ERROR!: XML to Word";
                MessageBox.Show(err.Message, err.StackTrace, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

}
