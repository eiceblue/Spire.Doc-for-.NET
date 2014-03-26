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

            //load a document
            document.LoadFromFile(@"..\..\..\..\..\..\Data\FindAndReplace.doc");

            //Find text
            TextSelection[] textSelections = document.FindAllString(this.textBox1.Text, true, true);

            //Set hightlight
            foreach(TextSelection selection in textSelections)
            {
                selection.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
            }

            //Save doc file.
            document.SaveToFile("Sample.doc", FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer("Sample.doc");
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
