using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace EnableTrackChanges
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
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_1.docx");

			// Enable tracking changes in the document.
			document.TrackChanges = true;

			// Specify the output file name for the modified document.
			String result = "Result-EnableTrackChanges.docx";

			// Save the document with enabled track changes to the specified file format (Docx2013).
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
