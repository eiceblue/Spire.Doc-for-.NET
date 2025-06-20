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

namespace GetFormFieldsCollection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			// Create a StringBuilder to hold the field information
			StringBuilder sb = new StringBuilder();

			// Load the document from a file
			Document document = new Document(@"..\..\..\..\..\..\Data\FillFormField.doc");

			// Get the first section of the document
			Section section = document.Sections[0];

			// Get the collection of form fields in the section
			FormFieldCollection formFields = section.Body.FormFields;

			// Append the count of form fields in the section to the StringBuilder
			sb.Append("The first section has " + formFields.Count + " form fields.");

			// Write the result to a text file
			File.WriteAllText("result.txt", sb.ToString());

			// Dispose the document object
			document.Dispose();

            //Launch result file
            WordDocViewer("result.txt");

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
