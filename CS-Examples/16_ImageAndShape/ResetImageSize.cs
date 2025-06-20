using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ResetImageSize
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\ImageTemplate.docx");

			//Get the first secion
			Section section = doc.Sections[0];

			//Get the first paragraph
			Paragraph paragraph = section.Paragraphs[0];

			//Reset the image size of the first paragraph
			foreach (DocumentObject docObj in paragraph.ChildObjects)
			{
				if (docObj is DocPicture)
				{
					DocPicture picture = docObj as DocPicture;

					//Set the width
					picture.Width = 50f;

					//Set the height
					picture.Height = 50f;
				}
			}

			//Save the document
			string output = "ResetImageSize.docx";
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
