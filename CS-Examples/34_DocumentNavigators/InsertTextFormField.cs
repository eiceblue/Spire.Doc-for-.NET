using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;


namespace InsertTextFormField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new empty Word document.
            Document doc = new Document();

            // Initialize a DocumentNavigator to help insert content and form fields into the document.
            DocumentNavigator navigator = new DocumentNavigator(doc);

            // Write a label indicating the type of the following form field: Calculation.
            navigator.Write("TextFormFieldType.Calculation: ");

            // Insert a calculation-type text form field.
            navigator.InsertTextFormField("CalculationTextField", TextFormFieldType.Calculation, "0", "=3+1", 30);

            // Insert a line break (carriage return) to move to the next line.
            navigator.Writeln();

            // Write a label for a number-only text form field.
            navigator.Write("TextFormFieldType.NumberText: ");

            // Insert a number-input text form field.
            navigator.InsertTextFormField("NumberText", TextFormFieldType.NumberText, "0", "100", 30);

            // Insert a line break.
            navigator.Writeln();

            // Write a label for a date-input text form field.
            navigator.Write("TextFormFieldType.DateText: ");

            // Insert a date-formatted text form field.
            navigator.InsertTextFormField("DateText", TextFormFieldType.DateText, "yyyy/M/d", "2025/8/1", 30);

            // Insert a line break.
            navigator.Writeln();

            // Enable automatic field updating. 
            doc.IsUpdateFields = true;

            // Save the resulting document to a file.
            doc.SaveToFile("InsertTextFormField.docx", FileFormat.Docx);

            // Close the document to release internal resources.
            doc.Close();

            // Explicitly dispose of the document object to free memory immediately.
            doc.Dispose();

            //Launch the MS Word file.
            WordDocViewer("InsertTextFormField.docx");
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
