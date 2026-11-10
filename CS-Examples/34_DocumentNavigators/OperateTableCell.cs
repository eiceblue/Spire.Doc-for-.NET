using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Formatting;

namespace OperateTableCell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new document instance.
            Document doc = new Document();

            // Initialize a document navigator to navigate and modify the document content.
            DocumentNavigator navigator = new DocumentNavigator(doc);

            // Begin creating a new table in the document.
            navigator.StartTable();

            // Insert the first cell into the current row of the table and get its reference.
            TableCell cell1 = navigator.InsertCell();

            // Add content to the first cell of the first row.
            addcellcontent(cell1, "Row 1, Cell 1");

            // Insert the second cell into the current row of the table and get its reference.
            TableCell cell2 = navigator.InsertCell();

            // Add content to the second cell of the first row.
            addcellcontent(cell2, "Row 1, Cell 2");

            // Insert the third cell into the current row of the table and get its reference.
            TableCell cell3 = navigator.InsertCell();

            // Add content to the third cell of the first row.
            addcellcontent(cell3, "Row 1, Cell 3");

            // End the current row and move to the next row in the table.
            navigator.EndRow();

            // Insert the first cell into the new row of the table and get its reference.
            TableCell cell4 = navigator.InsertCell();

            // Add content to the first cell of the second row.
            addcellcontent(cell4, "Row 2, Cell 1");

            // End the table creation process.
            navigator.EndTable();

            // Move the navigator's cursor to a specific cell in the first table
            navigator.MoveToCell(0, 0, 1, 0);

            // Insert (overwrite) the text at the current cursor position inside the target cell.
            navigator.Write("new content");

            // Get the formatting object of the current cell where the navigator is positioned.
            CellFormat cellformat = navigator.CellFormat;

            // Clear all existing formatting applied to the current cell.
            cellformat.ClearFormatting();

            // Set the background (shading) color of the cell to GreenYellow.
            cellformat.Shading.BackgroundPatternColor = Color.GreenYellow;

            // Save the modified document to the specified output file in DOCX format.
            doc.SaveToFile("OperateTableCell.docx", FileFormat.Docx);

            // Close the document to release internal resources.
            doc.Close();

            // Explicitly dispose of the document object to free memory immediately.
            doc.Dispose();

            //Launch the MS Word file.
            WordDocViewer("OperateTableCell.docx");
        }
        private void addcellcontent(TableCell cell, String Content)
        {
            // Add a new paragraph to the table cell.
            Paragraph para = cell.AddParagraph();

            // Append the specified text content to the paragraph inside the cell.
            para.AppendText(Content);

            // Center-align the text horizontally within the paragraph.
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            // Set the background (shading) color of the table cell to CadetBlue.
            cell.CellFormat.Shading.BackgroundPatternColor = Color.CadetBlue;

            // Vertically center the content within the table cell.
            cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
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
