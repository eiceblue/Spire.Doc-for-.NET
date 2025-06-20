using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Pages;
using Spire.Doc.Printing;

namespace PrintMultipageToOneSheet
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
            string inputFile = @"..\..\..\..\..\..\Data\Template_Docx_4.docx";

            // Create a new instance of Document
            Document doc = new Document();

			// Load the Word document from the specified input file
			doc.LoadFromFile(inputFile, FileFormat.Docx);

			// Create a new PrintDialog from System.Windows.Forms
			System.Windows.Forms.PrintDialog printDialog = new System.Windows.Forms.PrintDialog();

			// Enable printing to a file
			printDialog.PrinterSettings.PrintToFile = true;

			// Set the print file name based on the PagesPreSheet value
			printDialog.PrinterSettings.PrintFileName = string.Format("F:\\TP\\214\\sample-new2.xps");

			// Assign the PrintDialog to the document's PrintDialog
			doc.PrintDialog = printDialog;

			// Print the document with multiple pages condensed into one sheet
			doc.PrintMultipageToOneSheet(PagesPerSheet.FourPages, true);

            doc.Dispose();

            this.Close();

        }
    }
}
