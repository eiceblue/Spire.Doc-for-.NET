using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace LockedAspectRatio
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {

			// Create a new instance of Document
			Document document = new Document();

			// Add a new section to the document
			Section section = document.AddSection();

			// Add a paragraph to the section
			Paragraph paragraph = section.AddParagraph();

			// Append a textbox to the paragraph and get a reference to it
			Spire.Doc.Fields.TextBox textBox1 = paragraph.AppendTextBox(240, 35);

			// Configure the horizontal alignment, line color, and line style of the textbox
			textBox1.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
			textBox1.Format.LineColor = System.Drawing.Color.Black;
			textBox1.Format.LineStyle = TextBoxLineStyle.Simple;

			// Lock the aspect ratio of the textbox
			textBox1.AspectRatioLocked = true;

			// Add a paragraph to the body of the textbox and get a reference to it
			Paragraph para = textBox1.Body.AddParagraph();

			// Add text to the paragraph
			TextRange txtrg = para.AppendText("Textbox 1 in the document");
			txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
			txtrg.CharacterFormat.FontSize = 14;
			txtrg.CharacterFormat.TextColor = System.Drawing.Color.Black;
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			// Save the document to a file named "Sample.docx" in Docx format
			document.SaveToFile("Sample.docx", FileFormat.Docx);

			// Dispose the document object to release resources
			document.Dispose();

            //Launch the Word file.
            WordDocViewer("Sample.docx");
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
