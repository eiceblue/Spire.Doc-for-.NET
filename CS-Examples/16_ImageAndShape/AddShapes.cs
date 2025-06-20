using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AddShapes
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
			Section sec = doc.AddSection();

			// Add a paragraph to the section
			Paragraph para = sec.AddParagraph();

			int x = 60, y = 40, lineCount = 0;
			for (int i = 1; i < 20; i++)
			{
				// Check if the current line count is a multiple of 8
				if (lineCount > 0 && lineCount % 8 == 0)
				{
					// Append a page break to start a new page
					para.AppendBreak(BreakType.PageBreak);
					x = 60;
					y = 40;
					lineCount = 0;
				}

				// Append a shape to the paragraph
				ShapeObject shape = para.AppendShape(50, 50, (ShapeType)i);
				shape.HorizontalOrigin = HorizontalOrigin.Page;
				shape.HorizontalPosition = x;
				shape.VerticalOrigin = VerticalOrigin.Page;
				shape.VerticalPosition = y + 50;
				x = x + (int)shape.Width + 50;

				// Check if the shape count is a multiple of 5
				if (i > 0 && i % 5 == 0)
				{
					// Adjust the vertical position and line count
					y = y + (int)shape.Height + 120;
					lineCount++;
					x = 60;
				}
			}

			// Save the document
			doc.SaveToFile("AddShape.docx", FileFormat.Docx);

			// Dispose the document
			doc.Dispose();

            //Launch Word file.
            WordDocViewer("AddShape.docx");
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