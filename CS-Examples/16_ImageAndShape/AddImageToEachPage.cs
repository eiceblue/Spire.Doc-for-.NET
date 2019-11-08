using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Collections;
using System.Text;

namespace AddImageToEachPage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Open a Word document
            Document document = new Document(@"..\..\..\..\..\..\Data\SampleB_2.docx");

            string imgPath = @"..\..\..\..\..\..\Data\Spire.Doc.png";

            //Add a picture in footer and set it's position
            DocPicture picture = document.Sections[0].HeadersFooters.Footer.AddParagraph().AppendPicture(Image.FromFile(imgPath));
            picture.VerticalOrigin = VerticalOrigin.Page;
            picture.HorizontalOrigin = HorizontalOrigin.Page;
            picture.VerticalAlignment = ShapeVerticalAlignment.Bottom;
            picture.TextWrappingStyle = TextWrappingStyle.None;

            //Add a textbox in footer and set it's positiion
            Spire.Doc.Fields.TextBox textbox = document.Sections[0].HeadersFooters.Footer.AddParagraph().AppendTextBox(150, 20);
            textbox.VerticalOrigin = VerticalOrigin.Page;
            textbox.HorizontalOrigin = HorizontalOrigin.Page;
            textbox.HorizontalPosition = 300;
            textbox.VerticalPosition = 700;
            textbox.Body.AddParagraph().AppendText("Welcome to E-iceblue");
        
            //Save to file
            document.SaveToFile("result.docx", FileFormat.Docx);

            //Launch result file
            WordDocViewer("result.docx");

        }

        private void WordDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}
