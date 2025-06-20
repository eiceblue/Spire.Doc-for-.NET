using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;

namespace SetColumnWidth
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
         
			// Create a new document object
			Document document = new Document();

			// Load a document from a file, specified by the file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\TableSample.docx");

			// Get the first section of the document
			Section section = document.Sections[0];

			// Get the first table in the section
			Table table = section.Tables[0] as Table;

			// Set the width of the first column in each row to 200 points
			for (int i = 0; i < table.Rows.Count; i++)
			{
				table.Rows[i].Cells[0].SetCellWidth(200, CellWidthType.Point);
			}

			// Save the modified document to a file with the name "ColumnWidth.docx", using Docx format
			document.SaveToFile(@"ColumnWidth.docx", FileFormat.Docx);

			// Dispose of the document object
			document.Dispose();
			
            //Launch the document
            FileViewer("ColumnWidth.docx");
        }
        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
