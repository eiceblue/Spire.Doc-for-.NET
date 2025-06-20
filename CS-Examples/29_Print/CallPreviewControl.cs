using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing.Printing;

namespace CallPreviewControl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			// Specify the input file path
			String input = @"..\..\..\..\..\..\Data\Sample.docx";

			// Create a new instance of Document
			Document doc = new Document();

			// Load the Word document from the specified input file
			doc.LoadFromFile(input);

			// Get the PrintDocument associated with the document
			PrintDocument printDoc = doc.PrintDocument;

			// Create a new PrintPreviewDialog
			PrintPreviewDialog printPreviewDialog = new PrintPreviewDialog();

			// Set the PrintDocument for the PrintPreviewDialog
			printPreviewDialog.Document = doc.PrintDocument;

			// Set the size of the PrintPreviewDialog's client area
			printPreviewDialog.ClientSize = new Size(600, 800);

			// Show the PrintPreviewDialog
			printPreviewDialog.ShowDialog();

			// Dispose of the document object when finished using it
			doc.Dispose();
        }
    }
}
