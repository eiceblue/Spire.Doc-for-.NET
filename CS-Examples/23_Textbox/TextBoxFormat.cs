using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace TextBoxFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a new document
            Document doc = new Document();
            Section sec = doc.AddSection();

            //Add a text box and append sample text
            Spire.Doc.Fields.TextBox TB = doc.Sections[0].AddParagraph().AppendTextBox(310, 90);
            Paragraph para = TB.Body.AddParagraph();
            TextRange TR = para.AppendText("Using Spire.Doc, developers will find " +
                "a simple and effective method to endow their applications with rich MS Word features. ");
            TR.CharacterFormat.FontName = "Cambria ";
            TR.CharacterFormat.FontSize = 13;

            //Set exact position for the text box
            TB.Format.HorizontalOrigin = HorizontalOrigin.Page;
            TB.Format.HorizontalPosition = 120;
            TB.Format.VerticalOrigin = VerticalOrigin.Page;
            TB.Format.VerticalPosition = 100;

            //Set line style for the text box
            TB.Format.LineStyle = TextBoxLineStyle.Double;
            TB.Format.LineColor = Color.CornflowerBlue;
            TB.Format.LineDashing = LineDashing.Solid;
            TB.Format.LineWidth = 5;

            //Set internal margin for the text box
            TB.Format.InternalMargin.Top = 15;
            TB.Format.InternalMargin.Bottom = 10;
            TB.Format.InternalMargin.Left = 12;
            TB.Format.InternalMargin.Right = 10;

            //Save and launch document
            string output = "TextBoxFormat.docx";
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
