using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ReplaceTextWithTable
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

            // Find the first occurrence of the string "Christmas Day, December 25" in the document, and return its TextSelection object
            Section section = document.Sections[0];
            TextSelection selection = document.FindString("Christmas Day, December 25", true, true);

            // Convert the TextSelection object to a TextRange object
            TextRange range = selection.GetAsOneRange();
            // Get the Paragraph object that contains the TextRange object
            Paragraph paragraph = range.OwnerParagraph;
            // Get the text body that contains the paragraph
            Body body = paragraph.OwnerTextBody;
            // Find the index of the TextRange object in the ChildObjects collection of the Paragraph object
            int index = body.ChildObjects.IndexOf(paragraph);

            // Add a new table and reset the number of rows and columns to 3
            Table table = section.AddTable(true);
            table.ResetCells(3, 3);

            // Remove the TextRange object from the ChildObjects collection of the Paragraph object
            body.ChildObjects.Remove(paragraph);

            // Insert the table into the ChildObjects collection at the position of the previous TextRange object
            body.ChildObjects.Insert(index, table);

            // Define the output file path and filename
            string result = "Result-ReplaceTextWithTable.docx";

            // Save the document to the specified path with a .docx file format
            document.SaveToFile(result, FileFormat.Docx2013);

            // Dispose of the document object to release its resources
            document.Dispose();

            //Launch the MS Word file.
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
