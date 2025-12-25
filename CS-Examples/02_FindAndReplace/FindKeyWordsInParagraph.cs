using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace FindKeyWordsInParagraph
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Input file path
            String input = "..\\..\\..\\..\\..\\..\\Data\\Sample.docx";

            //Output file path
            String output = "FindKeyWordsInParagraph_output.docx";

            //Create word document
            Document document = new Document();

            //Load a document
            document.LoadFromFile(input);

            //Get the first section
            Section s = document.Sections[0];

            //Get the second paragraph
            Paragraph para = s.Paragraphs[1];

            //Find all matched keywords
            TextSelection[] textSelections = para.FindAllString("Word", false, true);

            //Highlight text
            foreach (TextSelection selection in textSelections)
            {
                selection.GetAsOneRange().CharacterFormat.HighlightColor= Color.FromArgb(255, 255, 0);
            }

            // Save to file
            document.SaveToFile(output, FileFormat.Docx2019);

            //Dispose the document
            document.Dispose();

            //Launching the Word file.
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
