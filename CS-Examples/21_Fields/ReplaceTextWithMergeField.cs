using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Collections;
using System.Text;

namespace ReplaceTextWithMergeField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
        
			// Load the document from a file
			Document document = new Document(@"..\..\..\..\..\..\Data\SampleB_2.docx");

			// Find the text "Test" in the document
			TextSelection ts = document.FindString("Test", true, true);

			// Get the selected text as a single range
			TextRange tr = ts.GetAsOneRange();

			// Get the paragraph that contains the selected text
			Paragraph par = tr.OwnerParagraph;

			// Get the index of the selected text within its parent paragraph
			int index = par.ChildObjects.IndexOf(tr);

			// Create a new merge field
			MergeField field = new MergeField(document);
			field.FieldName = "MergeField";

			// Insert the merge field at the same position as the selected text
			par.ChildObjects.Insert(index, field);

			// Remove the selected text from the paragraph
			par.ChildObjects.Remove(tr);

			// Save the modified document to a new file
			document.SaveToFile("result.docx", FileFormat.Docx);

			// Dispose of the document object
			document.Dispose();
			
            //Launch result file
            WordDocViewer("result.docx");

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
