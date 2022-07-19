using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ApplyEmphasisMark
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

            //Load the document from disk
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

            //Find text to emphasize
            TextSelection[] textSelections = document.FindAllString("Spire.Doc for .NET", false, true);

            //Set emphasis mark to the found text
            foreach (TextSelection selection in textSelections)
            {
                selection.GetAsOneRange().CharacterFormat.EmphasisMark = Emphasis.Dot;
            }

            //Save the file
            string output = "ApplyEmphasisMark.docx";
            document.SaveToFile(output, FileFormat.Docx);

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
