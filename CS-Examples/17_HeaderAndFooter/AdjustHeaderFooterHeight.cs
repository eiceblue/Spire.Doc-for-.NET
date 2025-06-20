using System;
using System.Windows.Forms;
using Spire.Doc;

namespace AdjustHeaderFooterHeight
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
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the first section
			Section section = doc.Sections[0];

			//Adjust the height of headers in the section
			section.PageSetup.HeaderDistance = 100;

			//Adjust the height of footers in the section
			section.PageSetup.FooterDistance = 100;

			//Save the document
			string output = "AdjustHeaderFooterHeight.docx";
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
