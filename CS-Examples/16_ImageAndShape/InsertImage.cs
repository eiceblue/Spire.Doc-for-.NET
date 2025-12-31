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
			string input = @"..\..\..\..\..\..\Data\BlankTemplate.docx";

			//Create a word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the first section
			Section section = doc.Sections[0];

			//Add a new section or get the first section
			Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

			//Append text
			paragraph.AppendText("The sample demonstrates how to insert an image into a document.");

			//Apply style
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			//Add a new paragraph
			paragraph = section.AddParagraph();

			//Append text
			paragraph.AppendText("The above is a picture.");

			//Load an image 
			Bitmap p = new Bitmap(Image.FromFile(@"..\..\..\..\..\..\Data\Word.png"));

			//rotate image and insert image to word document
			p.RotateFlip(RotateFlipType.Rotate90FlipX);

			//Create a DocPicture instance
			DocPicture picture = new DocPicture(doc);

			//Load the image
			picture.LoadImage(p);
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
             DocPicture picture = new DocPicture(doc);  
            picture.LoadImage(TestUtil.DataPath + "Demo/Word.png");

            */

            //set image's position
            picture.HorizontalPosition = 50.0F;
			picture.VerticalPosition = 60.0F;

			//set image's size
			picture.Width = 200;
			picture.Height = 200;

			//set textWrappingStyle with image;
			picture.TextWrappingStyle = TextWrappingStyle.Through;
			
			//Insert the picture at the beginning of the second paragraph
			paragraph.ChildObjects.Insert(0, picture);

			//Save the document
			string output = "InsertImageAtSpecifiedLocation.docx";
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the document
			doc.Dispose();
			
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
