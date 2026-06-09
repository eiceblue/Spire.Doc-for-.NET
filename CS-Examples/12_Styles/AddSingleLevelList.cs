using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddSingleLevelList
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class to represent a Word document
            Document document = new Document();

            // Add a new section to the document, which acts as a container for content like paragraphs and tables
            Section section = document.AddSection();

            // Define a list template that uses Arabic numerals (1, 2, 3) followed by a dot
            ListTemplate template = ListTemplate.NumberArabicDot;

            // Register this single-level numbered list template with the document and get a reference to it
            ListDefinitionReference listRef = document.ListReferences.AddSingleLevelList(template);

            // Create a new paragraph object within the current section
            Paragraph paragraph = section.AddParagraph();

            // Append the text to the newly created paragraph
            paragraph.AppendText("List Item 1");

            // Apply the previously defined numbered list format (listRef) at level 0 to this paragraph
            paragraph.ListFormat.ApplyListRef(listRef, 0);

            // Reassign the paragraph variable by adding another new paragraph to the section
            paragraph = section.AddParagraph();

            // Append the text to this new paragraph
            paragraph.AppendText("List Item 2");

            // Apply the same numbered list format at level 0 to continue the sequence
            paragraph.ListFormat.ApplyListRef(listRef, 0);

            // Create a third paragraph in the section for the next list item
            paragraph = section.AddParagraph();

            // Append the text to the paragraph
            paragraph.AppendText("List Item 3");

            // Apply the numbered list format at level 0 to complete the list
            paragraph.ListFormat.ApplyListRef(listRef, 0);

            string result = "addSingleLevelList.docx";
            // Save the document to a file using Docx format
            document.SaveToFile(result, FileFormat.Docx);

            // Close the document to release system resources associated with the file
            document.Close();

            // Dispose of the document object to free up memory
            document.Dispose();

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
