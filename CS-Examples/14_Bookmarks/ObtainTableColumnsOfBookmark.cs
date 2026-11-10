using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;

namespace ObtainTableColumnsOfBookmark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class.
            Document document = new Document();

            // Load a Word document from the specified relative file path.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ObtainTableColumnsOfBookmark.docx");

            // Retrieve the bookmark from the document.
            Bookmark bookmark = document.Bookmarks.FindByName("tableTest");

            // Get the index of the first column covered by the bookmark.
            int firstColumn = bookmark.FirstColumn;

            // Get the index of the last column covered by the bookmark.
            int lastColumn = bookmark.LastColumn;

            // Obtain the table cell that contains the start of the bookmark.
            TableCell tableCell_start = bookmark.BookmarkStart.OwnerParagraph.Owner as TableCell;

            // Get the table row that owns the starting cell.
            TableRow tableRow_start = tableCell_start.OwnerRow;

            // Get the table that contains the starting row.
            Table table = tableRow_start.Owner as Table;

            // Determine the row index of the starting row within the table.
            int rowIndex_start = table.Rows.IndexOf(tableRow_start);

            // Obtain the table cell that contains the end of the bookmark.
            TableCell tableCell_end = bookmark.BookmarkEnd.OwnerParagraph.Owner as TableCell;
            // Get the table row that owns the ending cell.
            TableRow tableRow_end = tableCell_end.OwnerRow;
            // Determine the row index of the ending row within the table.
            int rowIndex_end = table.Rows.IndexOf(tableRow_end);

            // Set the text of the starting cell (first column of the start row) to indicate bookmark start.
            table.Rows[rowIndex_start].Cells[firstColumn].FirstParagraph.Text = "Bookmark starts";
            // Highlight the starting cell with a yellow background.
            table.Rows[rowIndex_start].Cells[firstColumn].CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
            // Set the text of the ending cell (last column of the end row) to indicate bookmark completion.
            table.Rows[rowIndex_end].Cells[lastColumn].FirstParagraph.Text = "Bookmark completed";
            // Highlight the ending cell with a yellow background.
            table.Rows[rowIndex_end].Cells[lastColumn].CellFormat.Shading.BackgroundPatternColor = Color.Yellow;

            // Save the modified document to a new file in .docx format.
            document.SaveToFile("ObtainTableColumnsOfBookmark.docx", FileFormat.Docx);

            // Close the document to release resources.
            document.Close();

            // Explicitly dispose of the document object to free memory.
            document.Dispose();

            //Launching the Word file.
            WordDocViewer("ObtainTableColumnsOfBookmark.docx");
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
