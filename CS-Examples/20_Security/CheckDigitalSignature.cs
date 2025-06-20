using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;

namespace CheckDigitalSignature
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
                bool hasDigitalSignature = Document.HasDigitalSignature(@"..\..\..\..\..\..\Data\CheckDigitalSignature.docx");

				// Use a switch statement to determine the file format and update the fileFormat string accordingly
				if(hasDigitalSignature)
                {
                    MessageBox.Show("This Word document has digital signature");
                }else
                {
                    MessageBox.Show("This Word document has not digital signature");
                }
        }
    }
}
