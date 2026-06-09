using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddListTemplate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object to represent the Word document
            Document document = new Document();

            // Add a new section to the document, which is a fundamental structural element
            Section section = document.AddSection();

            // Define a default bullet list template for creating unordered lists
            ListTemplate template = ListTemplate.BulletDefault;

            // Register the bullet list template with the document and get a reference to it
            ListDefinitionReference listRef = document.ListReferences.Add(template);

            // Define a default numbered list template for creating ordered lists
            ListTemplate template1 = ListTemplate.NumberDefault;

            // Register the numbered list template with the document and get a reference to it
            ListDefinitionReference listRef1 = document.ListReferences.Add(template1);

            // Create a new paragraph within the current section
            Paragraph paragraph = section.AddParagraph();

            // Add the text "List Item 1" to the newly created paragraph
            paragraph.AppendText("List Item 1");

            // Apply the bullet list format (listRef) at level 0 to this paragraph
            paragraph.ListFormat.ApplyListRef(listRef, 0);

            // Create another new paragraph for the next list item
            paragraph = section.AddParagraph();

            // Add the text "List Item 2" to the paragraph
            paragraph.AppendText("List Item 2");

            // Apply the bullet list format at level 1 (a nested level) to this paragraph
            paragraph.ListFormat.ApplyListRef(listRef, 1);

            // Create a third paragraph for the final bullet point
            paragraph = section.AddParagraph();

            // Add the text "List Item 3" to the paragraph
            paragraph.AppendText("List Item 3");

            // Apply the bullet list format at level 2 (a deeper nested level) to this paragraph
            paragraph.ListFormat.ApplyListRef(listRef, 2);

            // Start the numbered list by creating a new paragraph
            paragraph = section.AddParagraph();

            // Add the text "List Item 6" to the paragraph
            paragraph.AppendText("List Item 6");

            // Apply the numbered list format (listRef1) at level 0 to this paragraph
            paragraph.ListFormat.ApplyListRef(listRef1, 0);

            // Create a new paragraph for the second numbered item
            paragraph = section.AddParagraph();

            // Add the text "List Item 7" to the paragraph
            paragraph.AppendText("List Item 7");

            // Apply the numbered list format at level 1 to this paragraph
            paragraph.ListFormat.ApplyListRef(listRef1, 1);

            // Create a new paragraph for the third numbered item
            paragraph = section.AddParagraph();

            // Add the text "List Item 8" to the paragraph
            paragraph.AppendText("List Item 8");

            // Apply the numbered list format at level 2 to this paragraph
            paragraph.ListFormat.ApplyListRef(listRef1, 2);

            string result = "AddTemplateList.docx";
            // Save the completed document to a file in Docx format
            document.SaveToFile(result, FileFormat.Docx);

            // Close the document, releasing any associated resources like file handles
            document.Close();

            // Dispose of the document object, freeing up system memory
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
