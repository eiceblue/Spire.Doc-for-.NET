using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace FindAndHighlight
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

            //Load the document from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

            //Find text
            TextSelection[] textSelections = document.FindAllString("word", false, true);

            //Set hightlight
            foreach(TextSelection selection in textSelections)
            {
                selection.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
            }

            //Save doc file.
            document.SaveToFile("Sample.docx", FileFormat.Docx);

            //Launching the  Word file.
            WordDocViewer("Sample.docx");
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
