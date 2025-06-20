using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ExtractParagraphBasedOnStyle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object.
            Document document = new Document();

            // Define the name of the style.
            String styleName1 = "Heading1";

            // Create a StringBuilder object to store the text with the specified style.
            StringBuilder style1Text = new StringBuilder();

            // Load the Word document file from the specified path.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractParagraphBasedOnStyle.docx");

            // Append a line to the StringBuilder indicating the style name.
            style1Text.AppendLine("The following is the content of the paragraph with the style name " + styleName1 + ": ");

            // Iterate over each section in the document.
            foreach (Section section in document.Sections)
            {
                // Iterate over each paragraph in the section.
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    // Check if the paragraph has the specified style.
                    if (paragraph.StyleName != null && paragraph.StyleName.Equals(styleName1))
                    {
                        // Append the text of the paragraph to the StringBuilder.
                        style1Text.AppendLine(paragraph.Text);
                    }
                }
            }

            // Define the output file name.
            string output1 = "ExtractParagraphBasedOnStyle_style1.txt";

            // Write the contents of the StringBuilder to the output file.
            File.WriteAllText(output1, style1Text.ToString());

            // Dispose the Document object to free up resources.
            document.Dispose();

            WordDocViewer(output1);
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
