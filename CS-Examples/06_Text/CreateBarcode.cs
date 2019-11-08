using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CreateBarcode
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a document
            Document doc = new Document();

            //Add a paragraph
            Paragraph p = doc.AddSection().AddParagraph();

            //Add barcode and set its format
            TextRange txtRang = p.AppendText("H63TWX11072");
            //Set barcode font name, note you need to install the barcode font on your system at first
            txtRang.CharacterFormat.FontName = "C39HrP60DlTt";
            txtRang.CharacterFormat.FontSize = 80;
            txtRang.CharacterFormat.TextColor = Color.SeaGreen;

            //Save and launch document
            string output = "CreateBarcode.docx";
            doc.SaveToFile(output, FileFormat.Docx);
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
