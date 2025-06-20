using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Reporting;

namespace AddHyperlinkForMailMergedImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object
            Document doc = new Document();
            // Load a Word document from a specific file path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\AddHyperlinkForImage.docx");
            // Define the field names and corresponding image file names
            var fieldNames = new string[] { "MyImage" };
            var fieldValues = new string[] { @"..\..\..\..\..\..\Data\mailmerge_logo.png" };
            // Attach an event handler for the MergeImageField event
            doc.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MailMerge_MergeImageField);
            // Execute the mail merge with the field names and values
            doc.MailMerge.Execute(fieldNames, fieldValues);
            // Save the modified document to a new file
            doc.SaveToFile("AddHyperlinkForImage.docx", FileFormat.Docx);
   

            WordDocViewer("AddHyperlinkForImage.docx");
        }

        // Event handler for the MergeImageField event
        private void MailMerge_MergeImageField(object sender, MergeImageFieldEventArgs field)
        {
            string filePath = field.ImageFileName;  // FieldValue as string;
            if (!string.IsNullOrEmpty(filePath))
            {
                field.Image = Image.FromFile(filePath);
                // Set the hyperlink for the merged image field
                field.ImageLink = "https://www.e-iceblue.com/";
            }
        }

        private void WordDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch(Exception e) {
                Debug.Write(e.StackTrace);
            }
        }

    }
}
