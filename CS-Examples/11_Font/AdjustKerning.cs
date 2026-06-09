using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AdjustKerning
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

            // Create a list to store test data (text description and kerning value)
            List<object[]> testData = new List<object[]>();

            // Add a test case for negative kerning
            testData.Add(new Object[] { "Negative Kerning (-1.0f)", -1.0f });

            // Add a test case for zero kerning (disables kerning)
            testData.Add(new Object[] { "Zero Kerning (0.0f)", 0.0f });

            // Add a test case for positive kerning
            testData.Add(new Object[] { "Positive Kerning (2.5f)", 2.5f });

            // Add a test case for a large kerning value
            testData.Add(new Object[] { "Huge Kerning (1638.0f)", 1638.0f });

            // Add a test case for a value exceeding the standard limit (1-1638)
            testData.Add(new Object[] { "Tiny Kerning (1639.0f)", 1639.0f });

            // Loop through each test data item
            foreach (object[] item in testData)
            {
                // Extract the text description from the first column
                String text = (string)item[0];

                // Extract the kerning value from the second column
                float kerningValue = (float)item[1];

                // Add a new paragraph to the section
                Paragraph pragraph = section.AddParagraph();

                // Append the text to the paragraph and get the text range
                TextRange textRange = pragraph.AppendText(text);

                // Apply the specific kerning value to the character format
                textRange.CharacterFormat.Kerning = kerningValue;
            }

            // Define the file name for the output document
            String result = "Adjust Kerning.docx";

            // Save the document to a file in Docx format
            doc.SaveToFile(result, FileFormat.Docx);

            // Close the document to release resources
            doc.Close();

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
