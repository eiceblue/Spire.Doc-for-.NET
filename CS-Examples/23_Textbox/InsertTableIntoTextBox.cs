using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertTableIntoTextBox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
        
			// Create a new document
			Document doc = new Document();

			// Add a section to the document
			Section section = doc.AddSection();

			// Add a paragraph to the section
			Paragraph paragraph = section.AddParagraph();

			// Append a text box to the paragraph with specified dimensions
			Spire.Doc.Fields.TextBox textbox = paragraph.AppendTextBox(300, 100);

			// Set the horizontal and vertical positioning of the text box
			textbox.Format.HorizontalOrigin = HorizontalOrigin.Page;
			textbox.Format.HorizontalPosition = 140;
			textbox.Format.VerticalOrigin = VerticalOrigin.Page;
			textbox.Format.VerticalPosition = 50;

			// Add a paragraph to the text box
			Paragraph textboxParagraph = textbox.Body.AddParagraph();

			// Append text to the paragraph in the text box
			TextRange textboxRange = textboxParagraph.AppendText("Table 1");
			textboxRange.CharacterFormat.FontName = "Arial";

			// Add a table to the body of the text box
			Table table = textbox.Body.AddTable(true);

			// Reset the number of rows and columns in the table
			table.ResetCells(4, 4);

			// Define the data for the table
			string[,] data = new string[,]
			{
				{"Name","Age","Gender","ID" },
				{"John","28","Male","0023" },
				{"Steve","30","Male","0024" },
				{"Lucy","26","female","0025" }
			};

			// Populate the table with data
			for (int i = 0; i < 4; i++)
			{
				for (int j = 0; j < 4; j++)
				{
					TextRange tableRange = table[i, j].AddParagraph().AppendText(data[i, j]);
					tableRange.CharacterFormat.FontName = "Arial";
				}
			}

			// Apply a predefined table style to the table
			table.ApplyStyle(DefaultTableStyle.TableColorful2);

			// Save the document to a file
			string output = "InsertTableIntoTextBox.docx";
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the document object
			doc.Dispose();
			
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
