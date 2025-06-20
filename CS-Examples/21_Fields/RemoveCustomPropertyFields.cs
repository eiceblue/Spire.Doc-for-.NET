using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemoveCustomPropertyFields
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
			// Create a new document object
			Document document = new Document();

			// Load an existing document from the specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveCustomPropertyFields.docx");

			// Get the collection of custom document properties
			CustomDocumentProperties cdp = document.CustomDocumentProperties;

			// Iterate through the custom document properties and remove them
			for (int i = 0; i < cdp.Count; )
			{
				cdp.Remove(cdp[i].Name);
			}

			// Enable the automatic update of fields in the document
			document.IsUpdateFields = true;

			// Specify the name for the resulting document file
			String result = "Result-RemoveCustomPropertyFields.docx";

			// Save the modified document to a file with the specified name and format
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose of the document object to free up resources
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
