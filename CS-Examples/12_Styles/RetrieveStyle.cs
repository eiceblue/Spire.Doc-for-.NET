using System;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RetrieveStyle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create and load a Word document
			Document doc = new Document(@"..\..\..\..\..\..\Data\Styles.docx");

			//Traverse all paragraphs in the document and get their style names through StyleName property
			string styleName = null;

			//Loop through all the sections
			foreach (Section section in doc.Sections)
			{
				//Loop through all the paragraphs
				foreach (Paragraph paragraph in section.Paragraphs)
				{
					//Get the style name
					styleName += paragraph.StyleName + "\r\n";
				}
			}

			//Save the text file
			string output = "RetrieveStyle.txt";
			File.WriteAllText(output, styleName.ToString());


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
