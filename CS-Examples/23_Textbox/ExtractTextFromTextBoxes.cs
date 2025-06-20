using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;
using Spire.Doc.Fields;
using System.Windows.Forms;

namespace ExtractTextFromTextBoxes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
      
			// Create a new Document object
			Document document = new Document();

			// Load the Word document from a file
			document.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractTextFromTextBoxes.docx");

			// Specify the output file name
			String result = "Result-ExtractTextFromTextBoxes.txt";

			// Check if the document contains any text boxes
			if (document.TextBoxes.Count > 0)
			{
				// Create a StreamWriter to write the extracted text to the output file
				using (StreamWriter sw = File.CreateText(result))
				{
					// Iterate through the sections in the document
					foreach (Section section in document.Sections)
					{
						// Iterate through the paragraphs in each section
						foreach (Paragraph p in section.Paragraphs)
						{
							// Iterate through the child objects of each paragraph
							foreach (DocumentObject obj in p.ChildObjects)
							{
								// Check if the child object is a text box
								if (obj.DocumentObjectType == DocumentObjectType.TextBox)
								{
									// Cast the child object to a TextBox
									Spire.Doc.Fields.TextBox textbox = obj as Spire.Doc.Fields.TextBox;

									// Iterate through the child objects of the text box
									foreach (DocumentObject objt in textbox.ChildObjects)
									{
										// Check if the child object is a paragraph
										if (objt.DocumentObjectType == DocumentObjectType.Paragraph)
										{
											// Write the text of the paragraph to the output file
											sw.Write((objt as Paragraph).Text);
										}

										// Check if the child object is a table
										if (objt.DocumentObjectType == DocumentObjectType.Table)
										{
											// Cast the child object to a Table
											Table table = objt as Table;

											// Extract text from the table and write it to the output file
											ExtractTextFromTables(table, sw);
										}
									}
								}
							}
						}
					}
				}
			}

			// Dispose the Document object to free up resources
			document.Dispose();

            //Launch the result file.
            WordDocViewer(result);
        }

      
		// Define a method to extract text from tables
		static void ExtractTextFromTables(Table table, StreamWriter sw)
		{
			// Iterate through the rows of the table
			for (int i = 0; i < table.Rows.Count; i++)
			{
				// Get the current row
				TableRow row = table.Rows[i];
				
				// Iterate through the cells of the row
				for (int j = 0; j < row.Cells.Count; j++)
				{
					// Get the current cell
					TableCell cell = row.Cells[j];

					// Iterate through the paragraphs in the cell
					foreach (Paragraph paragraph in cell.Paragraphs)
					{
						// Write the text of the paragraph to the output file
						sw.Write(paragraph.Text);
					}
				}
			}
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
