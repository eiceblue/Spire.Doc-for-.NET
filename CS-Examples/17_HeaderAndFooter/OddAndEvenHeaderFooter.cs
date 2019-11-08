using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace OddAndEvenHeaderFooter
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
            string input = @"..\..\..\..\..\..\Data\MultiplePages.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the section and
            Section section = doc.Sections[0];

            //Set the DifferentOddAndEvenPagesHeaderFooter property to ture
            section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = true;

            //Add odd header
            Paragraph P3 = section.HeadersFooters.OddHeader.AddParagraph();
            TextRange OH = P3.AppendText("Odd Header");
            P3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            OH.CharacterFormat.FontName = "Arial";
            OH.CharacterFormat.FontSize = 10;

            //Add even header
            Paragraph P4 = section.HeadersFooters.EvenHeader.AddParagraph();
            TextRange EH = P4.AppendText("Even Header from E-iceblue Using Spire.Doc");
            P4.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            EH.CharacterFormat.FontName = "Arial";
            EH.CharacterFormat.FontSize = 10;

            //Add odd footer
            Paragraph P2 = section.HeadersFooters.OddFooter.AddParagraph();
            TextRange OF = P2.AppendText("Odd Footer");
            P2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            OF.CharacterFormat.FontName = "Arial";
            OF.CharacterFormat.FontSize = 10;

            //Add even footer
            Paragraph P1 = section.HeadersFooters.EvenFooter.AddParagraph();
            TextRange EF = P1.AppendText("Even Footer from E-iceblue Using Spire.Doc");
            EF.CharacterFormat.FontName = "Arial";
            EF.CharacterFormat.FontSize = 10;
            P1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            //Save and launch document
            string output = "OddAndEvenHeaderFooter.docx";
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
