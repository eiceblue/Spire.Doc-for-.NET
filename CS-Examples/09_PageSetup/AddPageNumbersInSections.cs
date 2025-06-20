using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddPageNumbersInSections
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class
			Document document = new Document();

			// Load a Word document from a specific file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_4.docx");

			// Iterate through the first three sections of the document
			for (int i = 0; i < 3; i++)
			{
				// Get the footer of the current section
				HeaderFooter footer = document.Sections[i].HeadersFooters.Footer;

				// Add a paragraph to the footer
				Paragraph footerParagraph = footer.AddParagraph();

				// Append a page number field to the footer paragraph
				footerParagraph.AppendField("page number", FieldType.FieldPage);

				// Append " of " text to the footer paragraph
				footerParagraph.AppendText(" of ");

				// Append a section pages field to the footer paragraph
				footerParagraph.AppendField("number of pages", FieldType.FieldSectionPages);

				// Set the horizontal alignment of the footer paragraph to right
                footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

				// If it's the last iteration, exit the loop; otherwise, set up page numbering for the next section
				if (i == 2)
					break;
				else
				{
					// Restart page numbering for the next section
					document.Sections[i + 1].PageSetup.RestartPageNumbering = true;

					// Set the starting page number for the next section to 1
					document.Sections[i + 1].PageSetup.PageStartingNumber = 1;
				}
			}

			// Specify the file name for the resulting document with page numbers in sections
			string result = "Result-AddPageNumbersInSections.docx";

			// Save the modified document to the specified file path in the DOCX format (version: Word 2013)
			document.SaveToFile(result, FileFormat.Docx2013);

			// Release the resources used by the document object
			document.Dispose();

            //Launch the Ms Word file.
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
