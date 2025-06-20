using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertPageBreakFirstApproach
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
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_2.docx");

			// Find all occurrences of the word "technology" in the document
			TextSelection[] selections = document.FindAllString("technology", true, true);

			// Iterate through each found text selection
			foreach (TextSelection ts in selections)
			{
				// Get the range of the text selection as one continuous range
				TextRange range = ts.GetAsOneRange();

				// Get the paragraph that contains the text range
				Paragraph paragraph = range.OwnerParagraph;

				// Get the index of the text range within the paragraph's child objects
				int index = paragraph.ChildObjects.IndexOf(range);

				// Insert a page break after the text range by creating a Break object with BreakType.PageBreak
				Break pageBreak = new Break(document, BreakType.PageBreak);

				// Insert the page break at the next index position in the paragraph's child objects
				paragraph.ChildObjects.Insert(index + 1, pageBreak);
			}

			// Specify the file name for the resulting document with inserted page breaks
			string result = "Result-InsertPageBreakAtSpecifiedLocation.docx";

			// Save the modified document to the specified file path in the DOCX format (version: Word 2013)
			document.SaveToFile(result, FileFormat.Docx2013);

			// Release the resources used by the document object
			document.Dispose();

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
