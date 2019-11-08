using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddLineNumbers
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create Word document.
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

            //Set the start value of the line numbers.
            document.Sections[0].PageSetup.LineNumberingStartValue = 1;

            //Set the interval between displayed numbers.
            document.Sections[0].PageSetup.LineNumberingStep = 6;

            //Set the distance between line numbers and text.
            document.Sections[0].PageSetup.LineNumberingDistanceFromText = 40f;

            //Set the numbering mode of line numbers. There are four choices: None, Continuous, RestartPage and RestartSection.
            document.Sections[0].PageSetup.LineNumberingRestartMode = LineNumberingRestartMode.Continuous;
            
            String result = "Result-AddLineNumbers.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the MS Word file.
            WordDocViewer(result);
        }

        private void WordDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
