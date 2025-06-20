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
            string input = @"..\..\..\..\..\..\Data\MultiplePages.docx";

			//Create a word document
			Document doc = new Document();

			//Save to file
			doc.LoadFromFile(input);

			//Get the first section
			Section section = doc.Sections[0];

			//Set the DifferentOddAndEvenPagesHeaderFooter property to ture
			section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = true;

			//Add odd header
			Paragraph P3 = section.HeadersFooters.OddHeader.AddParagraph();

			//Append text
			TextRange OH = P3.AppendText("Odd Header");

			//Set the HorizontalAlignment for the paragraph
            P3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			//Set the font name and font size
			OH.CharacterFormat.FontName = "Arial";
			OH.CharacterFormat.FontSize = 10;

			//Add even header
			Paragraph P4 = section.HeadersFooters.EvenHeader.AddParagraph();

			//Append text
			TextRange EH = P4.AppendText("Even Header from E-iceblue Using Spire.Doc");

			//Set the HorizontalAlignment for the paragraph
            P4.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			//Set the font name and font size
			EH.CharacterFormat.FontName = "Arial";
			EH.CharacterFormat.FontSize = 10;

			//Add odd footer
			Paragraph P2 = section.HeadersFooters.OddFooter.AddParagraph();

			//Append text
			TextRange OF = P2.AppendText("Odd Footer");

			//Set the HorizontalAlignment for the paragraph
            P2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			//Set the font name and font size
			OF.CharacterFormat.FontName = "Arial";
			OF.CharacterFormat.FontSize = 10;

			//Add even footer
			Paragraph P1 = section.HeadersFooters.EvenFooter.AddParagraph();

			//Append text
			TextRange EF = P1.AppendText("Even Footer from E-iceblue Using Spire.Doc");

			//Set the font name and font size
			EF.CharacterFormat.FontName = "Arial";
			EF.CharacterFormat.FontSize = 10;

			//Set the HorizontalAlignment for the paragraph
            P1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			//Save the document
			string output = "OddAndEvenHeaderFooter.docx";
			doc.SaveToFile(output, FileFormat.Docx);

			//Dispose the document
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
