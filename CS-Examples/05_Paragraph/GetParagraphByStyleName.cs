using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;

namespace GetParagraphByStyleName
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

            // Load an existing Word document from the specified file path.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_DocX_3.docx");

            // Create a StringBuilder object to store the content.
            StringBuilder content = new StringBuilder();

            // Append a line of text to the StringBuilder.
            content.AppendLine("Get paragraphs by style name \"Heading1\": ");

            // Iterate over the Sections in the document.
            foreach (Section section in document.Sections)
            {
                // Iterate over the Paragraphs in each Section.
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    // Check if the Paragraph has the style name "Heading1".
                    if (paragraph.StyleName == "Heading1")
                    {
                        // Append the text of the Paragraph to the StringBuilder.
                        content.AppendLine(paragraph.Text);
                    }
                }
            }

            // Specify the file name for the resulting text file.
            String result = "Result-GetParagraphsByStyleName.txt";

            // Write the content of the StringBuilder to a text file.
            File.WriteAllText(result, content.ToString());

            // Dispose the Document object.
            document.Dispose();

            //Launch the file.
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
