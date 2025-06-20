using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertingTextbox
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

			// Call the method to insert a textbox into the section
			InsertTextbox(section);

			// Save the document to a file named "Sample.docx" in Docx format
			document.SaveToFile("Sample.docx", FileFormat.Docx);

			// Dispose the document object to release resources
			document.Dispose();

            //Launching the Word file.
            WordDocViewer("Sample.docx");


        }


		private void InsertTextbox(Section section)
		{
			// Create a paragraph in the specified section
			Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

			// Add three paragraphs to create space between textboxes
			paragraph = section.AddParagraph();
			paragraph = section.AddParagraph();
			paragraph = section.AddParagraph();

			// Create and customize textbox 1
			Spire.Doc.Fields.TextBox textBox1 = paragraph.AppendTextBox(240, 35);
			textBox1.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
			textBox1.Format.LineColor = System.Drawing.Color.Gray;
			textBox1.Format.LineStyle = TextBoxLineStyle.Simple;
			textBox1.Format.FillColor = System.Drawing.Color.DarkSeaGreen;
			Paragraph para = textBox1.Body.AddParagraph();
			TextRange txtrg = para.AppendText("Textbox 1 in the document");
			txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
			txtrg.CharacterFormat.FontSize = 14;
			txtrg.CharacterFormat.TextColor = System.Drawing.Color.White;
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			// Add four paragraphs to create space between textboxes
			paragraph = section.AddParagraph();
			paragraph = section.AddParagraph();
			paragraph = section.AddParagraph();
			paragraph = section.AddParagraph();

			// Create and customize textbox 2
			Spire.Doc.Fields.TextBox textBox2 = paragraph.AppendTextBox(240, 35);
			textBox2.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
			textBox2.Format.LineColor = System.Drawing.Color.Tomato;
			textBox2.Format.LineStyle = TextBoxLineStyle.ThinThick;
			textBox2.Format.FillColor = System.Drawing.Color.Blue;
			textBox2.Format.LineDashing = LineDashing.Dot;
			para = textBox2.Body.AddParagraph();
			txtrg = para.AppendText("Textbox 2 in the document");
			txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
			txtrg.CharacterFormat.FontSize = 14;
			txtrg.CharacterFormat.TextColor = System.Drawing.Color.Pink;
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			// Add four paragraphs to create space between textboxes
			paragraph = section.AddParagraph();
			paragraph = section.AddParagraph();
			paragraph = section.AddParagraph();
			paragraph = section.AddParagraph();

			// Create and customize textbox 3
			Spire.Doc.Fields.TextBox textBox3 = paragraph.AppendTextBox(240, 35);
			textBox3.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
			textBox3.Format.LineColor = System.Drawing.Color.Violet;
			textBox3.Format.LineStyle = TextBoxLineStyle.Triple;
			textBox3.Format.FillColor = System.Drawing.Color.Pink;
			textBox3.Format.LineDashing = LineDashing.DashDotDot;
			para = textBox3.Body.AddParagraph();
			txtrg = para.AppendText("Textbox 3 in the document");
			txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
			txtrg.CharacterFormat.FontSize = 14;
			txtrg.CharacterFormat.TextColor = System.Drawing.Color.Tomato;
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
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
