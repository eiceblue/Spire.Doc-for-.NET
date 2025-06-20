using Spire.Doc;
using System;
using System.Windows.Forms;

namespace PrintMultipleCopies
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
       
			// Create a new instance of Document
			Document document = new Document();

			// Load the Word document from the specified file
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

			// Set the printer name to "Microsoft Print to PDF" for printing
			document.PrintDocument.PrinterSettings.PrinterName = "Microsoft Print to PDF";

			// Set the number of copies to be printed to 4
			document.PrintDocument.PrinterSettings.Copies = 4;

			// Print the document
			document.PrintDocument.Print();

			// Dispose of the document object when finished using it
			document.Dispose();

        }

    }
}
