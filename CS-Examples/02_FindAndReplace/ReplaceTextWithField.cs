using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ReplaceTextWithField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Word document object
            Document document = new Document();

            // Load a Word document from a specific path
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceTextWithField.docx");

            // Find the first occurrence of the string "summary" in the document, and return its TextSelection object
            TextSelection selection = document.FindString("summary", false, true);

            // Convert the TextSelection object to a TextRange object
            TextRange textRange = selection.GetAsOneRange();

            // Get the Paragraph object that contains the TextRange object
            Paragraph ownParagraph = textRange.OwnerParagraph;

            // Find the index of the TextRange object in the ChildObjects collection of the Paragraph object
            int rangeIndex = ownParagraph.ChildObjects.IndexOf(textRange);

            // Remove the TextRange object from the ChildObjects collection of the Paragraph object at its index
            ownParagraph.ChildObjects.RemoveAt(rangeIndex);

            // Create a new list to store cloned objects
            List<DocumentObject> tempList = new List<DocumentObject>();

            // Loop through the ChildObjects collection of the Paragraph object, starting from the index after the removed TextRange object
            for (int i = rangeIndex; i < ownParagraph.ChildObjects.Count; i++)
            {
                // Clone the current object in the ChildObjects collection and add it to the tempList
                tempList.Add(ownParagraph.ChildObjects[rangeIndex].Clone());
                // Remove the current object from the ChildObjects collection at its index
                ownParagraph.ChildObjects.RemoveAt(rangeIndex);
            }

            // Append a field called "MyFieldName" to the end of the Paragraph, with a field type of MergeField
            ownParagraph.AppendField("MyFieldName", FieldType.FieldMergeField);

            // Loop through each object in the tempList
            foreach (DocumentObject obj in tempList)
            {
                // Add each object from the tempList back into the ChildObjects collection of the Paragraph
                ownParagraph.ChildObjects.Add(obj);
            }

            // Define the output file path and filename
            string output = "ReplaceTextWithField_output.docx";

            // Save the document to the specified path with a .docx file format
            document.SaveToFile(output, FileFormat.Docx);

            // Dispose of the document object to release its resources
            document.Dispose();
			
            WordDocViewer(output);
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
