using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemoveFooter
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

			//Create a Word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the first section
			Section section = doc.Sections[0];

			//Traverse the word document and clear all footers in different type
			foreach (Paragraph para in section.Paragraphs)
			{
				//Loop through the child objects of paragraph
				foreach (DocumentObject obj in para.ChildObjects)
				{
					//Clear footer in the first page
					HeaderFooter footer;
					footer = section.HeadersFooters[HeaderFooterType.FooterFirstPage];
					if (footer != null)
						footer.ChildObjects.Clear();
					//Clear footer in the odd page
					footer = section.HeadersFooters[HeaderFooterType.FooterOdd];
					if (footer != null)
						footer.ChildObjects.Clear();
					//Clear footer in the even page
					footer = section.HeadersFooters[HeaderFooterType.FooterEven];
					if (footer != null)
						footer.ChildObjects.Clear();
				}
			}

			//Save and launch document
			string output = "RemoveFooter.docx";
			doc.SaveToFile(output, FileFormat.Docx);

			//Dispose the document
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
