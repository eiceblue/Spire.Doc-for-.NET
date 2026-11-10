using Spire.Doc;
using Spire.Doc.Documents;
using System;
using System.Windows.Forms;

namespace DeleteCustomStyles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // Create a new Document instance.
            Document document = new Document();

            // Load an existing Word document that contains custom styles from the specified file path.
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\CustomStyles.docx");

            // Iterate through each section in the document.
            foreach (Spire.Doc.Section section in document.Sections)
            {
                // Iterate through each paragraph in the body of the current section.
                foreach (Paragraph paragraph in section.Body.Paragraphs)
                {
                    // Remove the style associated with the current paragraph from the document's style collection.
                    document.Styles[paragraph.StyleName].RemoveSelf();
                }

                // Iterate through each table in the body of the current section.
                foreach (Table table in section.Body.Tables)
                {
                    // Remove the style applied to the table itself.
                    document.Styles[table.Format.StyleName].RemoveSelf();

                    // Iterate through each row in the current table.
                    foreach (Spire.Doc.TableRow tableRow in table.Rows)
                    {
                        // Iterate through each cell in the current table row.
                        foreach (Spire.Doc.TableCell tableCell in tableRow.Cells)
                        {
                            // Iterate through each paragraph inside the current table cell.
                            foreach (Paragraph cellParagraph in tableCell.Paragraphs)
                            {
                                // Remove the style associated with the paragraph inside the table cell.
                                document.Styles[cellParagraph.StyleName].RemoveSelf();
                            }
                        }
                    }
                }
            }

            // Save the modified document to a new file with custom styles removed.
            document.SaveToFile("DeleteCustomStyles.docx", FileFormat.Docx);

            // Close the document to release internal resources.
            document.Close();

            // Explicitly dispose of the document object to free memory.
            document.Dispose();
            //Launching the Word file.
            WordDocViewer("DeleteCustomStyles.docx");
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
