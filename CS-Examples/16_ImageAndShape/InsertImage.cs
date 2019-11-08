using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertImage
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
            Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
            paragraph.AppendText("The sample demonstrates how to insert an image into a document.");
            paragraph.ApplyStyle(BuiltinStyle.Heading2);
            paragraph = section.AddParagraph();
            paragraph.AppendText("The above is a picture.");
            //get original image 
            Bitmap p = new Bitmap(Image.FromFile(@"..\..\..\..\..\..\Data\Word.png"));

            //rotate image and insert image to word document
            p.RotateFlip(RotateFlipType.Rotate90FlipX);

            //Create a picture
            DocPicture picture = new DocPicture(doc);
            picture.LoadImage(p);
            //set image's position
            picture.HorizontalPosition = 50.0F;
            picture.VerticalPosition = 60.0F;

            //set image's size
            picture.Width = 200;
            picture.Height = 200;

            //set textWrappingStyle with image;
            picture.TextWrappingStyle = TextWrappingStyle.Through;
            //Insert the picture at the beginning of the second paragraph
            paragraph.ChildObjects.Insert(0,picture);

            //Save and launch document
            string output = "InsertImageAtSpecifiedLocation.docx";
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
