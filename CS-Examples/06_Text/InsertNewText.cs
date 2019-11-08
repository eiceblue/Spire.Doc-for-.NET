using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertNewText
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
            string input = @"..\..\..\..\..\..\Data\Sample.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Find all the text string “New Zealand” from the sample document
            TextSelection[] selections = doc.FindAllString("Word", true, true);
            int index = 0;

            //Defines text range
            TextRange range = new TextRange(doc);

            //Insert new text string (New) after the searched text string
            foreach (TextSelection selection in selections)
            {
                range = selection.GetAsOneRange();
                TextRange newrange = new TextRange(doc);
                newrange.Text = ("(New text)");
                index = range.OwnerParagraph.ChildObjects.IndexOf(range);
                range.OwnerParagraph.ChildObjects.Insert(index + 1, newrange);
            }

            //Find and highlight the newly added text string New
            TextSelection[] text = doc.FindAllString("New text", true, true);
            foreach (TextSelection seletion in text)
            {
                seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
            }

            //Save and launch document
            string output = "InsertNewText.docx";
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
