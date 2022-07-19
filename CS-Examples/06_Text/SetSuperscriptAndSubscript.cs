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
            //Create word document
            Document document = new Document();

            //Create a new section
            Section section = document.AddSection();

            Paragraph paragraph = section.AddParagraph();
            paragraph.AppendText("E = mc");
            TextRange range1 = paragraph.AppendText("2");

            //Set supperscript
            range1.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

            paragraph.AppendBreak(BreakType.LineBreak);
            paragraph.AppendText("F");
            TextRange range2 = paragraph.AppendText("n");

            //Set subscript
            range2.CharacterFormat.SubSuperScript = SubSuperScript.SubScript;

            paragraph.AppendText(" = F");
            paragraph.AppendText("n-1").CharacterFormat.SubSuperScript = SubSuperScript.SubScript;
            paragraph.AppendText(" + F");
            paragraph.AppendText("n-2").CharacterFormat.SubSuperScript = SubSuperScript.SubScript;

            //Set font size
            foreach (var i in paragraph.Items)
            {
                if (i is TextRange)
                {
                    (i as TextRange).CharacterFormat.FontSize = 36;
                }
            }

            //Save the file
            string output = "SetSuperscriptAndSubscript.docx";
            document.SaveToFile(output,FileFormat.Docx);

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
