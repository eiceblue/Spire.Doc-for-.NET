using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;

namespace PrintDocViaXpsPrint
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                //Create a document
                using (Document document = new Document())
                {
                    //Load file
                    document.LoadFromFile(@"..\..\..\..\..\..\Data\Template.docx");
                    //Save to Xps file
                    document.SaveToStream(ms, FileFormat.XPS);
                }
                ms.Position = 0;
                String printerName = "HP LaserJet P1007";
                XpsPrint.XpsPrintHelper.Print(ms, printerName, "My printing job", true);
            }
        }
    }
}
