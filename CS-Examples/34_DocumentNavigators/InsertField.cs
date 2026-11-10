using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace InsertField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new DocumentNavigator instance, which automatically creates an underlying empty document.
            DocumentNavigator navigator = new DocumentNavigator();

            // Get the Document object associated with the navigator.
            Document doc = navigator.Document;

            // Retrieve the current (and initially only) section of the document.
            Section section = navigator.CurrentSection;

            // Insert a new paragraph at the current position and assign it to 'par'.
            Paragraph par = navigator.InsertParagraph();

            // Write the label text "Add page fields:" followed by a paragraph break.
            navigator.Writeln("Add page fields:");

            // Insert a PAGE field with numeric formatting.
            navigator.InsertField("PAGE \\# \"Page 0\"");

            // Insert a paragraph break (empty line).
            navigator.Writeln();

            // Insert the same PAGE field but with a custom result text to simulate a placeholder or initial value.
            navigator.InsertField("PAGE \\# \"Page 0\"", "3");

            // Insert another paragraph break.
            navigator.Writeln();

            // Insert a built-in PAGE field using FieldType enumeration, with the field result displayed (true = show result).
            navigator.InsertField(FieldType.FieldPage, true);

            // Insert another paragraph break.
            navigator.Writeln();

            // Insert another PAGE field, but this time hide the field result (false = show field code instead of result).
            navigator.InsertField(FieldType.FieldPage, false);

            // Save the document to a file in DOCX format.
            doc.SaveToFile("InsertField.docx", FileFormat.Docx);


            // Close the document to release internal resources.
            doc.Close();

            // Explicitly dispose of the document object to free memory immediately.
            doc.Dispose();

            //Launch the MS Word file.
            WordDocViewer("InsertField.docx");
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
