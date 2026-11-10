using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Interface;

namespace AddCustomStyles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of a Word document.
            Document doc = new Document();

            // Add a new section to the document.
            Section section = doc.AddSection();

            // Add a new paragraph to the section.
            Paragraph paragraph = section.AddParagraph();

            // Append initial text to the paragraph and obtain a reference to the resulting text range.
            TextRange textRange = paragraph.AppendText("Spire.Doc for .NET ");

            // Create and register a new character-style named "MyCharacterStyle1", then cast it to ICharacterStyle.
            ICharacterStyle characterStyle = (ICharacterStyle)doc.Styles.Add(StyleType.CharacterStyle, "MyCharacterStyle1");

            // Set the font name of the character style to Arial.
            characterStyle.CharacterFormat.FontName = "Arial";

            // Set the font size of the character style to 14 points.
            characterStyle.CharacterFormat.FontSize = 14;

            // Enable bold formatting for the character style.
            characterStyle.CharacterFormat.Bold = true;

            // Set the text color of the character style to blue.
            characterStyle.CharacterFormat.TextColor = Color.Blue;

            // Apply the custom character style to the previously created text range.
            textRange.ApplyStyle(characterStyle.Name);

            // Append additional descriptive text to the same paragraph.
            paragraph.AppendText("is a proven reliable MS Word API for .NET which enables to perform many Word document processing tasks.");

            // Create and register a new paragraph-style named "MyParagraphStyle1", then cast it to ParagraphStyle.
            ParagraphStyle heading1Style = (ParagraphStyle)doc.Styles.Add(StyleType.ParagraphStyle, "MyParagraphStyle1");

            // Set the paragraph alignment to justified (text aligned evenly along both left and right margins).
            heading1Style.ParagraphFormat.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Justify;

            // Set the line spacing of the paragraph style to 18 points.
            heading1Style.ParagraphFormat.LineSpacing = 18;

            // Set the font name for the paragraph style.
            heading1Style.CharacterFormat.FontName = "Calibri";

            // Set the font size for the paragraph style.
            heading1Style.CharacterFormat.FontSize = 12;

            // Set a custom text color using RGB values.
            heading1Style.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136);

            // Add another new paragraph to the section.
            paragraph = section.AddParagraph();
            paragraph.AppendText("As an independent Word .NET API, Spire.Doc for .NET doesn't need Microsoft Word to be installed on neither the development nor target systems. However, it can incorporate Microsoft Word document creation capabilities into any developers' .NET applications.");

            // Apply the custom paragraph style to the previously created paragraph.
            paragraph.ApplyStyle(heading1Style);

            // Add a table to the section.
            Table table = section.AddTable();
            // Initialize the table with 15 rows and 4 columns.
            table.ResetCells(15, 4);
            // Loop through each row of the table.
            for (int i = 0; i < 15; i++)
            {
                // Get the current row.
                TableRow row = table.Rows[i];
                // Loop through each cell in the current row.
                for (int j = 0; j < 4; j++)
                {
                    // Get the current cell.
                    TableCell cell = row.Cells[j];
                    // Add a paragraph to the cell and insert placeholder text.
                    // The text alternates between "Start" and "Continuation" based on column index.
                    cell.AddParagraph().AppendText(string.Format("Row {0}", (j % 4 == 0 ? "Start" : "Continuation")));
                }
            }

            // Create a custom table style named "MyTableStyle1".
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.TableStyle, "MyTableStyle1");
            // Set the border color for the table.
            tableStyle.Borders.Color = Color.Black;
            // Set the border type to double line.
            tableStyle.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Double;
            // Apply row striping every 3 rows.
            tableStyle.RowStripe = 3;
            // Set background color for odd-numbered striped rows.
            tableStyle.ConditionalStyles[TableConditionalStyleType.OddRowStripe].Shading.BackgroundPatternColor = Color.LightBlue;
            // Set background color for even-numbered striped rows.
            tableStyle.ConditionalStyles[TableConditionalStyleType.EvenRowStripe].Shading.BackgroundPatternColor = Color.LightCyan;

            // Apply column striping every 1 column (i.e., every other column).
            tableStyle.ColumnStripe = 1;
            // Set background color for even-numbered striped columns.
            tableStyle.ConditionalStyles[TableConditionalStyleType.EvenColumnStripe].Shading.BackgroundPatternColor = Color.LightYellow;

            // Apply the custom table style to the table.
            table.ApplyStyle(tableStyle);

            // Enable column striping in the table's formatting options.
            table.Format.StyleOptions = table.Format.StyleOptions | TableStyleOptions.ColumnStripe;

            // Save the document to a file in .docx format.
            doc.SaveToFile("AddCustomStyles.docx", FileFormat.Docx);

            // Close the document to release resources.
            doc.Close();
            // Explicitly dispose of the document object.
            doc.Dispose();

            //Launching the Word file.
            WordDocViewer("AddCustomStyles.docx");
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
