using System;
using System.Windows.Forms;
using Spire.Doc;
using System.Drawing.Printing;

namespace CustomPaperSize
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

            // Set the paper size of the default page settings to a custom size
            printDoc.DefaultPageSettings.PaperSize = new PaperSize("custom", 900, 800);

            // Print the document
            printDoc.Print();

            // Dispose of the document object when finished using it
            doc.Dispose();
        }
    }
}
