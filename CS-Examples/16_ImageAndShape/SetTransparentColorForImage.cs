using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetTransparentColorForImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\ImageTemplate.docx";

			//Create a word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the first paragraph in the first section
			Paragraph paragraph = doc.Sections[0].Paragraphs[0];

			//Loop through the child objects of the paragraph
			foreach (DocumentObject obj in paragraph.ChildObjects)
			{
				if (obj is DocPicture)
				{
					//Set the blue color of the image(s) in the paragraph to transperant
					DocPicture picture = obj as DocPicture;
					picture.TransparentColor = Color.Blue;
				}
			}

			//Save the document
			string output = "SetTransparentColorForImage.docx";
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
