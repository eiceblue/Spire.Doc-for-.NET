using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;
using Spire.Doc.Fields;

namespace ReplaceImageWithText
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

			//Replace all pictures with texts
			int j = 1;
			foreach (Section sec in doc.Sections)
			{
				foreach (Paragraph para in sec.Paragraphs)
				{
					List<DocumentObject> pictures = new List<DocumentObject>();
					//Get all pictures in the Word document
					foreach (DocumentObject docObj in para.ChildObjects)
					{
						if (docObj.DocumentObjectType == DocumentObjectType.Picture)
						{
							pictures.Add(docObj);
						}
					}

					//Replace pitures with the text "Here was image {image index}"
					foreach (DocumentObject pic in pictures)
					{
						//Get the index of the picture
						int index = para.ChildObjects.IndexOf(pic);

						//Create a new TextRange
						TextRange range = new TextRange(doc);

						//Format the text
						range.Text = string.Format("Here was image {0}", j);

						//Insert the textrange
						para.ChildObjects.Insert(index, range);

						//Remove the picture
						para.ChildObjects.Remove(pic);
						j++;
					}
				}
			}

			//Save and launch document
			string output = "ReplaceWithTexts.docx";
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
