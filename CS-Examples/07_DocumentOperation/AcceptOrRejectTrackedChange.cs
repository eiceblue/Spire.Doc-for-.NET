using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AcceptOrRejectTrackedChange
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object.
			Document document = new Document();

			// Load a Word document from the specified file path.
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\AcceptOrRejectTrackedChanges.docx");

			// Get the first Section of the document.
			Section sec = document.Sections[0];

			// Get the first Paragraph of the Section.
			Paragraph para = sec.Paragraphs[0];

			// Accept all changes made in the document.
			para.Document.AcceptChanges();

			// Specify the output file name for the modified document.
			String result = "Result-AcceptOrRejectTrackedChanges.docx";

			// Save the document with accepted changes to the specified file format (Docx2013).
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose of the Document object to free up resources.
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
