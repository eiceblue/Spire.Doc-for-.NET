using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Documents;
using Spire.Doc.Interface;

namespace ModifyCustomStyles
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

            // Load an existing Word document that contains custom styles from the specified relative file path.
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\CustomStyles.docx");

            // Attempt to find a paragraph style named "MyParagraphStyle1" in the document's style collection.
            IStyle style1 = document.Styles.FindByName("MyParagraphStyle1");

            // Check if the style exists and is of type ParagraphStyle.
            if (style1 != null && style1 is ParagraphStyle)
            {
                // Cast the found style to ParagraphStyle for modification.
                ParagraphStyle cStyle = style1 as ParagraphStyle;

                // Set the font name of the paragraph style to "Arial".
                cStyle.CharacterFormat.FontName = "Arial";

                // Set the font size to 18 points.
                cStyle.CharacterFormat.FontSize = 18;

                // Disable bold formatting.
                cStyle.CharacterFormat.Bold = false;

                // Change the text color to gray.
                cStyle.CharacterFormat.TextColor = Color.Gray;
            }

            // Attempt to find a table style named "MyTableStyle1" in the document's style collection.
            IStyle style2 = document.Styles.FindByName("MyTableStyle1");

            // Check if the style exists and is of type TableStyle.
            if (style2 != null && style2 is TableStyle)
            {
                // Cast the found style to TableStyle for modification.
                TableStyle tableStyle = style2 as TableStyle;

                // Set the border color of the table to gray.
                tableStyle.Borders.Color = Color.Gray;

                // Change the border type to single line.
                tableStyle.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;

                // Set the border line width to 1 point.
                tableStyle.Borders.LineWidth = 1;

                // Apply row striping every 2 rows (i.e., stripe pattern repeats every two rows).
                tableStyle.RowStripe = 2;
            }

            // Save the modified document to a new file using the Word 2016 .docx format.
            document.SaveToFile("ModifyCustomStyles.docx", FileFormat.Docx2016);

            // Close the document to release associated resources.
            document.Close();

            // Explicitly dispose of the document object to free memory immediately.
            document.Dispose();

            //Launching the Word file.
            WordDocViewer("ModifyCustomStyles.docx");
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
