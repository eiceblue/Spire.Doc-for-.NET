using System;
using System.Windows.Forms;
using Spire.Doc;

namespace ToXLSX
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

            // Load an existing Word document
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ConvertedToXLSX.docx");

            // Define the file path and name for the output document
            String result = "ToXLSX.xlsx";

            // Convert the Word document to XLSX file
            document.SaveToFile(result, FileFormat.XLSX);

            // Close the document to release resources
            document.Close();

            // Dispose of the document object to free up memory
            document.Dispose();

            //Launching the Word file.
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
