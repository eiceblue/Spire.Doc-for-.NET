using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetEditableRange
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

			// Load the Word document file from the specified path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\SetEditableRange.docx");

			// Set the document protection to allow only reading with a password
			document.Protect(ProtectionType.AllowOnlyReading, "password");

			// Create a PermissionStart object to mark the start of an editable range with a specific ID
			PermissionStart start = new PermissionStart(document, "testID");
			// Create a PermissionEnd object to mark the end of the editable range with the same ID
			PermissionEnd end = new PermissionEnd(document, "testID");

			// Insert the PermissionStart object at the beginning of the first paragraph in the first section
			document.Sections[0].Paragraphs[0].ChildObjects.Insert(0, start);
			// Add the PermissionEnd object to the end of the first paragraph in the first section
			document.Sections[0].Paragraphs[0].ChildObjects.Add(end);

			// Specify the output file name for the modified document
			string output = "SetEditableRange_output.docx";

			// Save the modified document to the specified file format
			document.SaveToFile(output, FileFormat.Docx);

			// Dispose the Document object to free resources
			document.Dispose();
			
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
