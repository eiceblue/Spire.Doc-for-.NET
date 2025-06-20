using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetHyperlinkFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			// Specify the input file path for the document
			string input = @"..\..\..\..\..\..\Data\BlankTemplate.docx";

			// Create a new Document object
			Document doc = new Document();

			// Load the document from the specified file path
			doc.LoadFromFile(input);

			// Get the first section of the document
			Section section = doc.Sections[0];

			// Add a paragraph to the section and append regular text
			Paragraph para1 = section.AddParagraph();
			para1.AppendText("Regular Link: ");

			// Append a hyperlink to the paragraph with the specified URL and display text
			TextRange txtRange1 = para1.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);
			txtRange1.CharacterFormat.FontName = "Times New Roman";
			txtRange1.CharacterFormat.FontSize = 12;

			// Add a blank paragraph as separation
			Paragraph blankPara1 = section.AddParagraph();

			// Add another paragraph to the section and append text
			Paragraph para2 = section.AddParagraph();
			para2.AppendText("Change Color: ");

			// Append a hyperlink to the paragraph with the specified URL and display text
			TextRange txtRange2 = para2.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);
			txtRange2.CharacterFormat.FontName = "Times New Roman";
			txtRange2.CharacterFormat.FontSize = 12;
			txtRange2.CharacterFormat.TextColor = Color.Red;

			// Add a blank paragraph as separation
			Paragraph blankPara2 = section.AddParagraph();

			// Add another paragraph to the section and append text
			Paragraph para3 = section.AddParagraph();
			para3.AppendText("Remove Underline: ");

			// Append a hyperlink to the paragraph with the specified URL and display text
			TextRange txtRange3 = para3.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);
			txtRange3.CharacterFormat.FontName = "Times New Roman";
			txtRange3.CharacterFormat.FontSize = 12;
			txtRange3.CharacterFormat.UnderlineStyle = UnderlineStyle.None;

			// Specify the output file path for the modified document
			string output = "HyperlinkFormat.docx";

			// Save the modified document to the output file path in DOCX format
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the document object to free up resources
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
