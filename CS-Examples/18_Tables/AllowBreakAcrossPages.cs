using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AllowBreakAcrossPages
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create word document
			Document document = new Document();

			//Load the file from disk
			document.LoadFromFile(@"..\..\..\..\..\..\Data\AllowBreakAcrossPages.docx");

			//Get the first section
			Section section = document.Sections[0];

			//Get the first table
			Table table = section.Tables[0] as Table;

			//Loop through the table rows
			foreach (TableRow row in table.Rows)
			{
				//Allow break across pages
				row.RowFormat.IsBreakAcrossPages = true;
			}

			//Save the Word document
			string output = "AllowBreakAcrossPages_out.docx";
			document.SaveToFile(output, FileFormat.Docx2013);

			//Dispose the document
			document.Dispose();

            //Launch the file
            WordDocViewer(output);
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
