using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Text.RegularExpressions;

namespace ReplaceContentWithDoc
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
            Document document1 = new Document();
            // Load a Word document from a file
            document1.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceContentWithDoc.docx");
            // Create another instance of the Document class
            Document document2 = new Document();
            // Load another Word document from a file
            document2.LoadFromFile(@"..\..\..\..\..\..\Data\Insert.docx");
            // Get the first section of the first document
            Section section1 = document1.Sections[0];
            // Create a regular expression object to search for a pattern
            Regex regex = new Regex(@"\[MY_DOCUMENT\]", RegexOptions.None);
            // Find all occurrences of the pattern in the first document
            TextSelection[] textSections = document1.FindAllPattern(regex);
            // Loop through each occurrence of the pattern
            foreach (TextSelection seletion in textSections)
            {
                // Get the paragraph that contains the pattern
                Paragraph para = seletion.GetAsOneRange().OwnerParagraph;
                // Get the range of text that contains the pattern
                TextRange textRange = seletion.GetAsOneRange();
                // Get the index of the paragraph in the first document's section
                int index = section1.Body.ChildObjects.IndexOf(para);
                // Loop through each section in the second document
                foreach (Section section2 in document2.Sections)
                {
                    // Loop through each paragraph in the second section
                    foreach (Paragraph paragraph in section2.Paragraphs)
                    {
                        // Insert the paragraph from the second document into the first document's section
                        section1.Body.ChildObjects.Insert(index, paragraph.Clone() as Paragraph);
                    }
                }
                // Remove the range of text that contains the pattern from the paragraph
                para.ChildObjects.Remove(textRange);
            }
            // Save the modified first document to a file
            document1.SaveToFile("Output.docx", FileFormat.Docx);
            // Dispose the first document and release all resources it is using
            document1.Dispose();
            // Dispose the second document and release all resources it is using
            document2.Dispose();

            //Launch the Word file.
            WordDocViewer("Output.docx");
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
