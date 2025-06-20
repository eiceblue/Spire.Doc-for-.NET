using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace LockSpecifiedSections
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

			// Add two sections to the document
			Section s1 = document.AddSection();
			Section s2 = document.AddSection();

			// Add a paragraph with text to section 1
			s1.AddParagraph().AppendText("Spire.Doc demo, section 1");

			// Add a paragraph with text to section 2
			s2.AddParagraph().AppendText("Spire.Doc demo, section 2");

			// Protect the document with a password and allow only form fields
			document.Protect(ProtectionType.AllowOnlyFormFields, "123");

			// Disable form field protection for section 2
			s2.ProtectForm = false;

			// Specify the output file path for the locked document
			string result = "Result-LockSpecifiedSections.docx";

			// Save the locked document to the output file path in DOCX format (compatible with Word 2013)
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose the document object to free up resources
			document.Dispose();

            //Launch the file.
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
