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

            //Add the date field
            Field eqField = new Field(document);
            eqField.Type = FieldType.FieldDate;
            eqField.Code = @"DATE  \@" + "\"yyyyMMdd \"";
            ownParagraph.ChildObjects.Insert(rangeIndex, eqField);
            FieldMark mark = new FieldMark(document, FieldMarkType.FieldSeparator);
            ownParagraph.ChildObjects.Insert(rangeIndex+1, mark);  
            FieldMark end = new FieldMark(document, FieldMarkType.FieldEnd);
            ownParagraph.ChildObjects.Insert(rangeIndex  + 2, end);

            // Remove the selected text from the paragraph
            ownParagraph.ChildObjects.Remove(textRange);

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
