using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace InsertImageIntoTextBox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
      
			// Create a new Document object
			Document doc = new Document();

			// Add a section to the document
			Section section = doc.AddSection();

			// Add a paragraph to the section
			Paragraph paragraph = section.AddParagraph();

			// Append a text box to the paragraph with specified dimensions
			Spire.Doc.Fields.TextBox tb = paragraph.AppendTextBox(220, 220);

			// Set the horizontal and vertical positioning of the text box
			tb.Format.HorizontalOrigin = HorizontalOrigin.Page;
			tb.Format.HorizontalPosition = 50;
			tb.Format.VerticalOrigin = VerticalOrigin.Page;
			tb.Format.VerticalPosition = 50;

			// Set the background fill effect of the text box to a picture
			tb.Format.FillEfects.Type = BackgroundType.Picture;

			// Set the picture for the background fill effect from a file
			tb.Format.FillEfects.Picture = Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png");
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            tb.Format.FillEfects.Picture = (@"..\..\..\..\..\..\Data\Spire.Doc.png");
            */

            // Specify the output file name
            string output = "InsertImageIntoTextBox.docx";

			// Save the document to a file in DOCX format
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the Document object to free up resources
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
