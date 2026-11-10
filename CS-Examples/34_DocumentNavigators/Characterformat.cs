using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace Characterformat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // Load the document by creating a new Document instance.
            Document doc = new Document();

            // Create a DocumentNavigator object to facilitate easy content insertion and formatting.
            DocumentNavigator navigator = new DocumentNavigator(doc);

            // Write plain text into the document (without special formatting yet).
            navigator.Write("Write plain text into the document (without special formatting yet).");

            // Insert a line break after the previous text.
            navigator.InsertBreak(BreakType.LineBreak);

            // Enable underline formatting for subsequent text.
            navigator.CharacterFormat.UnderlineColor = Color.OrangeRed;
            navigator.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

            // Set bold formatting for subsequent text.
            navigator.CharacterFormat.Bold = true;

            // Enable shadow effect for subsequent text.
            navigator.CharacterFormat.IsShadow = true;

            // Set the text color to blue for subsequent text.
            navigator.CharacterFormat.TextColor = Color.Blue;

            // Write formatted text using the current character formatting settings.
            navigator.Write("Write formatted text using the current character formatting settings.");

            // Insert another line break.
            navigator.InsertBreak(BreakType.LineBreak);

            // Save the current character formatting settings onto an internal stack for later reuse.
            navigator.PushCharacterFormat(); 

            // Clear all character formatting to default (e.g., no bold, no color, etc.).
            navigator.CharacterFormat.ClearFormatting();

            // Write text with cleared (default) formatting.
            navigator.Write("Write text with cleared (default) formatting");

            // Insert another line break.
            navigator.InsertBreak(BreakType.LineBreak);

            // Restore the previously saved character formatting from the stack.
            navigator.PopCharacterFormat(); 

            // Write text using the restored formatting.
            navigator.Write("Write text using the restored formatting.");

            // Save the document to a file in DOCX format.
            doc.SaveToFile("Characterformat.docx", FileFormat.Docx);

            // Close the document to release internal resources.
            doc.Close();

            // Explicitly dispose of the document object to free memory immediately.
            doc.Dispose();

            //Launch the MS Word file.
            WordDocViewer("Characterformat.docx");
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
