using System;
using System.Windows.Forms;
using Spire.Doc;

namespace SetPositionAndNumberFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Specify the input file path
            String input = @"..\..\..\..\..\..\Data\Footnote.docx";

            // Create a new instance of Document
            Document doc = new Document();

            // Load the Word document from the specified input file
            doc.LoadFromFile(input);

            // Get the first section of the document
            Section sec = doc.Sections[0];

            // Set the footnote options for the section
            sec.FootnoteOptions.NumberFormat = FootnoteNumberFormat.UpperCaseLetter;
            sec.FootnoteOptions.RestartRule = FootnoteRestartRule.RestartPage;
            sec.FootnoteOptions.Position = FootnotePosition.PrintAsEndOfSection;

            // Specify the output file path
            String output = "SetPositionAndNumberFormat.docx";

            // Save the modified document to the specified output file
            doc.SaveToFile(output, FileFormat.Docx);

            // Dispose of the document object when finished using it
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
