using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Collections;

namespace RemoveField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
    
			// Load the document from a specified file path
			Document document = new Document(@"..\..\..\..\..\..\Data\IfFieldSample.docx");

			// Get the first field in the document
			Field field = document.Fields[0];

			// Get the parent paragraph of the field
			Paragraph par = field.OwnerParagraph;

			// Get the index of the field within the child objects of the paragraph
			int index = par.ChildObjects.IndexOf(field);

			// Remove the field from the paragraph
			par.ChildObjects.RemoveAt(index);

			// Save the modified document to a file with the specified name and format
			document.SaveToFile("result.docx", FileFormat.Docx);

			// Dispose of the document object to free up resources
			document.Dispose();

            //Launch the Word file
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
