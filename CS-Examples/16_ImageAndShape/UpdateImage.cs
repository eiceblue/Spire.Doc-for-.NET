using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;
using Spire.Doc.Fields;

namespace UpdateImage
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

			//Create a list to store the pictures
			List<DocumentObject> pictures = new List<DocumentObject>();

			//Loop through the sections
			foreach (Section sec in doc.Sections)
			{
				//Loop through the paragraphs
				foreach (Paragraph para in sec.Paragraphs)
				{

					//Loop through the child objects of the paragraph
					foreach (DocumentObject docObj in para.ChildObjects)
					{
						//Determine if the type is picture or not
						if (docObj.DocumentObjectType == DocumentObjectType.Picture)
						{
							//Add the picure to list
							pictures.Add(docObj);
						}
					}
				}
			}

			//Create a DocPicture instance
			DocPicture picture = pictures[0] as DocPicture;

			//Replace the first picture with a new image file
			picture.LoadImage(Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png"));
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
              picture.LoadImage(TestUtil.DataPath + "Demo/E-iceblue.png");
            */

            //Save the document
            string output = "ReplaceWithNewImage.docx";
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
