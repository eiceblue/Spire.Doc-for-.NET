using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace EmbedPrivateFont
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\BlankTemplate.docx";

			//Create a Word document.
			Document doc = new Document();

			//Load the file from disk.
			doc.LoadFromFile(input);

			//Get the first section
			Section section = doc.Sections[0];

			//Add a paragraph
			Paragraph p = section.AddParagraph();

			//Append text to the paragraph
			TextRange range = p.AppendText("Spire.Doc for .NET is a professional Word.NET library specifically designed for developers to create, read, write, convert and print Word document files from any.NET platform with fast and high quality performance.");

			//Set the font name
			range.CharacterFormat.FontName = "PT Serif Caption";

			//Set the font size
			range.CharacterFormat.FontSize = 20;

			//Allow embedding font in document
			doc.EmbedFontsInFile = true;

			//Embed private font from font file into the document
			doc.PrivateFontList.Add(new PrivateFontPath("PT Serif Caption", @"..\..\..\..\..\..\Data\PT Serif Caption.ttf"));

			string output = "EmbedPrivateFont.docx";
			
			//Save the document
			doc.SaveToFile(output, FileFormat.Docx);

			//Dispose the Document
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
