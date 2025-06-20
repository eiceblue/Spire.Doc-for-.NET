using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CopyHeaderAndFooter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\HeaderAndFooter.docx";

			//Create a word document
			Document doc1 = new Document();

			//Load the source file
			doc1.LoadFromFile(input);

			//Get the header section from the source document
			HeaderFooter header = doc1.Sections[0].HeadersFooters.Header;

			input = @"..\..\..\..\..\..\Data\Template.docx";
			
			//Create a word document
			Document doc2 = new Document();

			//Load the destination file
			doc2.LoadFromFile(input);

			//Loop through the sections of doc2
			foreach (Section section in doc2.Sections)
			{
				//Loop through the child objects of heder
				foreach (DocumentObject obj in header.ChildObjects)
				{
					//Copy each object in the header of source file to destination file
					section.HeadersFooters.Header.ChildObjects.Add(obj.Clone());
				}
			}

			//Save the document
			string output = "CopyHeaderAndFooter.docx";
			doc2.SaveToFile(output, FileFormat.Docx);

			// Dispose the documents
			doc1.Dispose();
			doc2.Dispose();
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
