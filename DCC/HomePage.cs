using System.Security.Principal;
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
            CheckUserLogin();

        }


        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.Text = "DCC CREATER";
            tabControl1.SelectedTab = HomeTab;
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
                // Feedback i�lemi
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

        private void CheckUserLogin()
        {
            // Windows kimlik nesnesi olu�tur
            WindowsIdentity identity = WindowsIdentity.GetCurrent();

            // Kullan�c�n�n ad�n� al
            string userName = identity.Name;

            // Bilgisayar oturumuna giri� yapan kullan�c�n�n ad�n� al
            string user = userName.Split('\\')[1];


            string hosgeldiniz = "Welcome";
            // Ho�geldin mesaj�n� g�ster
            label1.Text = hosgeldiniz;
            label2.Text = user;
        }

        private void HomePage_Load(object sender, EventArgs e)
        {
        }

        private void CertificatePageShowButton_Click_1(object sender, EventArgs e)
        {
            this.Visible = false;

            try
            {

               

                // Feedback
                labelProgress.Visible = false;
                Thread.Sleep(10);
                progressBar.Value = 0;
                for (int i = 0; i < 100; i++) progressBar.Value += 1;
                labelProgress.Visible = true;
                labelProgress.ForeColor = System.Drawing.Color.Green;
                labelProgress.Text = @"Administrative data save successful";
                CertificateForm certificateForm = new CertificateForm();
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

        private void XMLtoWordPageButton_Click_1(object sender, EventArgs e)
        {
            groupBox1.Visible = true;
            this.Text = "XML TO WORD";
            tabControl1.SelectedTab = XMLtoWordPage;
        }

  
    }

}
