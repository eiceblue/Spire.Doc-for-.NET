using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Collections.Generic;

namespace SetTextWrap
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

			//Loop through the sections
			foreach (Section sec in doc.Sections)
			{
				//Loop through the paragraphs
				foreach (Paragraph para in sec.Paragraphs)
				{
					//Create a list to store the pictures
					List<DocumentObject> pictures = new List<DocumentObject>();

					//Get all pictures in the Word document
					foreach (DocumentObject docObj in para.ChildObjects)
					{
						if (docObj.DocumentObjectType == DocumentObjectType.Picture)
						{
							pictures.Add(docObj);
						}
					}

					//Set text wrap styles for each piture
					foreach (DocumentObject pic in pictures)
					{
						//Create a DocPicture instance
						DocPicture picture = pic as DocPicture;

						//Set teh wrap style and type
						picture.TextWrappingStyle = TextWrappingStyle.Through;
						picture.TextWrappingType = TextWrappingType.Both;
					}
				}
			}

			//Save the document
			string output = "SetTextWrap.docx";
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
