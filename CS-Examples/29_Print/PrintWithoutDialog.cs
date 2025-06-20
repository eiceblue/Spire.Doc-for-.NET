using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing.Printing;

namespace PrintWithoutDialog
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

            // Set the print controller to StandardPrintController for silent printing
            printDoc.PrintController = new StandardPrintController();

            // Print the document
            printDoc.Print();

            // Dispose of the document object when finished using it
            doc.Dispose();
        }
    }
}
