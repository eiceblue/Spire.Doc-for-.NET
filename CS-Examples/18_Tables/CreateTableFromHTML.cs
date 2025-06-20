using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using System.IO;

namespace CreateTableFromHTML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
      
			// Define an HTML string representing a table structure
			String HTML = "<table border='2px'>" +
						  "<tr>" +
						  "<td>Row 1, Cell 1</td>" +
						  "<td>Row 1, Cell 2</td>" +
						  "</tr>" +
						  "<tr>" +
						  "<td>Row 2, Cell 1</td>" +
						  "<td>Row 2, Cell 2</td>" +
						  "</tr>" +
						  "</table>";

			// Create a new Document object
			Document document = new Document();

			// Add a new section to the document
			Section section = document.AddSection();

			// Append the HTML content to the section as a paragraph
			section.AddParagraph().AppendHTML(HTML);

			// Save the document to a file in Docx2013 format
			string output = "CreateTableFromHTML_out.docx";
			document.SaveToFile(output, FileFormat.Docx2013);

			// Dispose of the document object to free up resources
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
