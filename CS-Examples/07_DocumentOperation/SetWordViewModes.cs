using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetWordViewModes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class
			Document document = new Document();

			// Load a Word document from the specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

			// Set the document view type to WebLayout
			document.ViewSetup.DocumentViewType = DocumentViewType.WebLayout;

			// Set the zoom percentage to 150%
			document.ViewSetup.ZoomPercent = 150;

			// Set the zoom type to None
			document.ViewSetup.ZoomType = ZoomType.None;

			// Specify the file name for the result document
			String result = "Result-SetWordViewModes.docx";

			// Save the modified document to the specified file path in the Docx2013 format
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose of the document object to release resources
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
