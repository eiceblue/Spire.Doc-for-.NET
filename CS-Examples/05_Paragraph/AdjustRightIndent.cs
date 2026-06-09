using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AdjustRightIndent
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
            Document doc = new Document();

            // Add a new section to the document
            Section section = doc.AddSection();

            // Add a new paragraph to the body of the section
            Paragraph paragraph = section.Body.AddParagraph();

            // Set the text content of the paragraph
            paragraph.Text = "Hello World!";

            // Enable the adjustment of the right indent for the paragraph format
            paragraph.Format.AdjustRightIndent = true;


            // Add another new paragraph to the body of the section
            paragraph = section.Body.AddParagraph();

            // Set the text content for the second paragraph
            paragraph.Text = "Thank you for using the Spire.Doc product.";

            // Disable the adjustment of the right indent for this paragraph
            paragraph.Format.AdjustRightIndent = false;

            // Define the file path and name for the output document
            String result = "AdjustRightIndent.docx";

            // Save the document to a file in Docx 2016 format
            doc.SaveToFile(result, FileFormat.Docx2016);

            // Close the document to release resources
            doc.Close();

            // Dispose of the document object to free up memory
            doc.Dispose();


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
