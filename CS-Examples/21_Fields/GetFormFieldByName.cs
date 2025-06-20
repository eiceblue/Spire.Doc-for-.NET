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

namespace GetFormFieldByName
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

			// Get the form field with the name "email"
			FormField formField = section.Body.FormFields["email"];

			// Append the name and type of the form field to the StringBuilder
			sb.AppendLine("The name of the form field is " + formField.Name);
			sb.AppendLine("The type of the form field is " + formField.FormFieldType);

			// Write the field information to a text file
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
