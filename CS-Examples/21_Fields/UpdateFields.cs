using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Collections;

namespace UpdateFields
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
			Document document = new Document(@"..\..\..\..\..\..\Data\IfFieldSample.docx");

            		// Setting the culture source when updating fields
            		document.FieldOptions.CultureSource = Spire.Doc.Layout.Fields.FieldCultureSource.CurrentThread;

			// Enable automatic update of fields in the document
			document.IsUpdateFields = true;

			// Save the document to a new file
			document.SaveToFile("result.docx", FileFormat.Docx);

			// Dispose of the document object
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
