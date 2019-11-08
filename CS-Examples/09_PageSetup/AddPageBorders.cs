using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddPageBorders
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

            //Get the first section.
            Section section = document.Sections[0];

            //Add page borders with special style and color.
            section.PageSetup.Borders.BorderType = Spire.Doc.Documents.BorderStyle.DoubleWave;
            section.PageSetup.Borders.Color = Color.LightSeaGreen;

            //Set the space between border and text.
            section.PageSetup.Borders.Left.Space = 50;
            section.PageSetup.Borders.Right.Space = 50;

            String result = "Result-AddPageBorders.docx";

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
