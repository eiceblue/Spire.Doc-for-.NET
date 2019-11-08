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
            //Load the document
            string input = @"..\..\..\..\..\..\Data\BlankTemplate.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the first section and add a paragraph
            Section section = doc.Sections[0];
            Paragraph p = section.AddParagraph();

            //Append text to the paragraph, then set the font name and font size
            TextRange range = p.AppendText("Spire.Doc for .NET is a professional Word.NET library specifically designed for developers to create, read, write, convert and print Word document files from any.NET platform with fast and high quality performance.");
            range.CharacterFormat.FontName = "PT Serif Caption";
            range.CharacterFormat.FontSize = 20;

            //Allow embedding font in document
            doc.EmbedFontsInFile = true;

            //Embed private font from font file into the document
            doc.PrivateFontList.Add(new PrivateFontPath("PT Serif Caption", @"..\..\..\..\..\..\Data\PT Serif Caption.ttf"));

            //Save and launch document
            string output = "EmbedPrivateFont.docx";
            doc.SaveToFile(output, FileFormat.Docx);
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
