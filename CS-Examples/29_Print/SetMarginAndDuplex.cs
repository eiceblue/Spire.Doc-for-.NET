using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing.Printing;

namespace SetMarginAndDuplex
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

            // Set the OriginAtMargins property to true to align the printable area with the margins
            printDoc.OriginAtMargins = true;

            // Set the Margins property of the DefaultPageSettings to zero to remove any margins
            printDoc.DefaultPageSettings.Margins = new System.Drawing.Printing.Margins(0, 0, 0, 0);

            // Set the Duplex property of PrinterSettings to Vertical for double-sided printing
            printDoc.PrinterSettings.Duplex = Duplex.Vertical;

            // Print the document
            printDoc.Print();

            // Dispose of the document object when finished using it
            doc.Dispose();
        }
    }
}
