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
            //Load Document
            string input = @"..\..\..\..\..\..\Data\BlankTemplate.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);
            Section section = doc.Sections[0];

            //Add a paragraph and append a hyperlink to the paragraph
            Paragraph para1 = section.AddParagraph();
            para1.AppendText("Regular Link: ");
            //Format the hyperlink with default color and underline style
            TextRange txtRange1 = para1.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);
            txtRange1.CharacterFormat.FontName = "Times New Roman";
            txtRange1.CharacterFormat.FontSize = 12;
            Paragraph blankPara1 = section.AddParagraph();

            //Add a paragraph and append a hyperlink to the paragraph
            Paragraph para2 = section.AddParagraph();
            para2.AppendText("Change Color: ");
            //Format the hyperlink with red color and underline style
            TextRange txtRange2 = para2.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);
            txtRange2.CharacterFormat.FontName = "Times New Roman";
            txtRange2.CharacterFormat.FontSize = 12;
            txtRange2.CharacterFormat.TextColor = Color.Red;
            Paragraph blankPara2 = section.AddParagraph();

            //Add a paragraph and append a hyperlink to the paragraph
            Paragraph para3 = section.AddParagraph();
            para3.AppendText("Remove Underline: ");
            //Format the hyperlink with red color and no underline style
            TextRange txtRange3 = para3.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);
            txtRange3.CharacterFormat.FontName = "Times New Roman";
            txtRange3.CharacterFormat.FontSize = 12;
            txtRange3.CharacterFormat.UnderlineStyle = UnderlineStyle.None;

            //Save and launch document
            string output = "HyperlinkFormat.docx";
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
