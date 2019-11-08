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
            //Load the document
            string input = @"..\..\..\..\..\..\Data\Footnote.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the first section
            Section sec = doc.Sections[0];

            //Set the number format, restart rule and position for the footnote
            sec.FootnoteOptions.NumberFormat = FootnoteNumberFormat.UpperCaseLetter;
            sec.FootnoteOptions.RestartRule = FootnoteRestartRule.RestartPage;
            sec.FootnoteOptions.Position = FootnotePosition.PrintAsEndOfSection;

            //Save and launch document
            string output = "SetPositionAndNumberFormat.docx";
            doc.SaveToFile(output, FileFormat.Docx);
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
