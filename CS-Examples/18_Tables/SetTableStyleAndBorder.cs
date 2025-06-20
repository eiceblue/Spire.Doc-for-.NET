using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetTableStyleAndBorder
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

			// Apply a predefined table style to the table
			table.ApplyStyle(DefaultTableStyle.ColorfulList);

			// Set the right border of the table
			table.Format.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.Hairline;
			table.Format.Borders.Right.LineWidth = 1.0F;
			table.Format.Borders.Right.Color = Color.Red;

			// Set the top border of the table
			table.Format.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Hairline;
			table.Format.Borders.Top.LineWidth = 1.0F;
			table.Format.Borders.Top.Color = Color.Green;

			// Set the left border of the table
			table.Format.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.Hairline;
			table.Format.Borders.Left.LineWidth = 1.0F;
			table.Format.Borders.Left.Color = Color.Yellow;

			// Set the bottom border of the table
			table.Format.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.DotDash;

			// Set the vertical borders of the table
			table.Format.Borders.Vertical.BorderType = Spire.Doc.Documents.BorderStyle.Dot;
			table.Format.Borders.Vertical.Color = Color.Orange;

			// Set the horizontal borders of the table to none
			table.Format.Borders.Horizontal.BorderType = Spire.Doc.Documents.BorderStyle.None;

			// Save the modified document to a file named "TableStyleAndBorder.docx", using Docx format
			document.SaveToFile("TableStyleAndBorder.docx", FileFormat.Docx);

			// Dispose of the document object
			document.Dispose();
			
            FileViewer("TableStyleAndBorder.docx");
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
