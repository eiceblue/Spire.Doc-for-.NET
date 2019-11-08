using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CreateImageHyperlink
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load Document
            string input = @"..\..\..\..\..\..\Data\BlankTemplate.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            Section section = doc.Sections[0];
            //Add a paragraph
            Paragraph paragraph = section.AddParagraph();
            //Load an image to a DocPicture object
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png");
            DocPicture picture = new DocPicture(doc);
            //Add an image hyperlink to the paragraph
            picture.LoadImage(image);
            paragraph.AppendHyperlink("https://www.e-iceblue.com/Introduce/word-for-net-introduce.html", picture, HyperlinkType.WebLink);

            //Save and launch document
            string output = "CreateImageHyperlink.docx";
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
