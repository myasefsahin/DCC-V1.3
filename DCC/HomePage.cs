using System.Xml;

namespace DCC
{
    public partial class HomePage : Form
    {
       
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
    }
}
