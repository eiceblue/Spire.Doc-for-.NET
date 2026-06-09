using System;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Fields;
using System.IO;
using Spire.Doc.Documents;
using Spire.Doc.Settings;

namespace CompatibilityOptions
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize a new Document object
            Document doc = new Document();

            // Add a new section to the document
            Section section = doc.AddSection();

            // Add a new paragraph to the section
            Paragraph paragraph = section.AddParagraph();

            // Define a string containing a label and trailing spaces for the underline effect
            string blanks = "(6)                  ";

            // Append the text string to the paragraph and get the TextRange object
            TextRange tr = paragraph.AppendText(blanks);

            // Set the underline style of the text to Single
            tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

            // Enable compatibility option: Include trailing spaces in underlined text
            doc.CompatibilityOptions.UlTrailSpace = true;

            // Enable compatibility option: Adjust line height specifically for tables
            doc.CompatibilityOptions.AdjustLineHeightInTable = true;

            // Enable compatibility option: Reserve space for underline characters to prevent clipping
            doc.CompatibilityOptions.SpaceForUL = true;

            // Enable compatibility option: Apply complex script breaking rules for line breaks
            doc.CompatibilityOptions.ApplyBreakingRules = true;

            // Disable compatibility option: Allow expansion of lines ending with Shift+Return (manual line break)
            doc.CompatibilityOptions.DoNotExpandShiftReturn = false;

            // Disable compatibility option: Do not override font size and justification defined in table styles
            doc.CompatibilityOptions.OverrideTableStyleFontSizeAndJustification = false;

            // Enable compatibility option: Prevent automatic fitting of tables that have fixed width constraints
            doc.CompatibilityOptions.DoNotAutofitConstrainedTables = true;

            // Optimize the document's compatibility settings specifically for Word 2016
            doc.CompatibilityOptions.OptimizeForWordVersion(WordVersion.Word2016);

            // Define the output file name
            String outputFile = "CompatibilityOptions.docx";

            // Save the document to the specified file in DOCX 2019 format
            doc.SaveToFile(outputFile, FileFormat.Docx2019);

            // Close the document to release file handles
            doc.Close();
        
            // Dispose of the document object to free up memory
            doc.Dispose();

            WordDocViewer(outputFile);
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
