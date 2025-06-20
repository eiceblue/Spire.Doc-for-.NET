using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace TextBoxFormat
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
			Document doc = new Document();

			// Add a new section to the document
			Section sec = doc.AddSection();

			// Add a textbox to the first paragraph in the section and get a reference to it
			Spire.Doc.Fields.TextBox TB = doc.Sections[0].AddParagraph().AppendTextBox(310, 90);

			// Add a paragraph to the body of the textbox and get a reference to it
			Paragraph para = TB.Body.AddParagraph();

			// Add text to the paragraph
			TextRange TR = para.AppendText("Using Spire.Doc, developers will find " +
				"a simple and effective method to endow their applications with rich MS Word features. ");

			// Set the font properties for the text
			TR.CharacterFormat.FontName = "Cambria";
			TR.CharacterFormat.FontSize = 13;

			// Configure the position of the textbox
			TB.Format.HorizontalOrigin = HorizontalOrigin.Page;
			TB.Format.HorizontalPosition = 120;
			TB.Format.VerticalOrigin = VerticalOrigin.Page;
			TB.Format.VerticalPosition = 100;

			// Configure the line style and color of the textbox
			TB.Format.LineStyle = TextBoxLineStyle.Double;
			TB.Format.LineColor = Color.CornflowerBlue;
			TB.Format.LineDashing = LineDashing.Solid;
			TB.Format.LineWidth = 5;

			// Configure the internal margins of the textbox
			TB.Format.InternalMargin.Top = 15;
			TB.Format.InternalMargin.Bottom = 10;
			TB.Format.InternalMargin.Left = 12;
			TB.Format.InternalMargin.Right = 10;

			// Specify the output file path
			string output = "TextBoxFormat.docx";

			// Save the modified document to the output file with the specified file format (Docx)
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the document object to release resources
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
