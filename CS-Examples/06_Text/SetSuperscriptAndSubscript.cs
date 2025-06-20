using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetSuperscriptAndSubscript
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class.
			Document document = new Document();

			// Add a section to the document.
			Section section = document.AddSection();

			// Add a paragraph to the section.
			Paragraph paragraph = section.AddParagraph();

			// Append the text "E = mc" to the paragraph.
			paragraph.AppendText("E = mc");

			// Append the text "2" as a superscript to the paragraph.
			TextRange range1 = paragraph.AppendText("2");
			range1.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

			// Insert a line break in the paragraph.
			paragraph.AppendBreak(BreakType.LineBreak);

			// Append the text "F" to the paragraph.
			paragraph.AppendText("F");

			// Append the text "n" as a subscript to the paragraph.
			TextRange range2 = paragraph.AppendText("n");
			range2.CharacterFormat.SubSuperScript = SubSuperScript.SubScript;

			// Append the text " = Fn-1 + Fn-2" with specific subscripts to the paragraph.
			paragraph.AppendText(" = F");
			paragraph.AppendText("n-1").CharacterFormat.SubSuperScript = SubSuperScript.SubScript;
			paragraph.AppendText(" + F");
			paragraph.AppendText("n-2").CharacterFormat.SubSuperScript = SubSuperScript.SubScript;

			// Set the font size to 36 for all TextRange items in the paragraph.
			foreach (var i in paragraph.Items)
			{
				if (i is TextRange)
				{
					(i as TextRange).CharacterFormat.FontSize = 36;
				}
			}

			// Specify the output file name.
			string output = "SetSuperscriptAndSubscript.docx";

			// Save the document to a file with the specified output file name and format (Docx).
			document.SaveToFile(output, FileFormat.Docx);

			// Clean up resources used by the document.
			document.Dispose();

            //Launching the file
            WordDocViewer(output);

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
