using Spire.Doc;
using System;
using System.Windows.Forms;


namespace RemoveAndReplaceComment
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
            string input = @"..\..\..\..\..\..\Data\CommentSample.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Replace the content of the first comment
            doc.Comments[0].Body.Paragraphs[0].Replace("This is the title", "This comment is changed.", false, false);

            //Remove the second comment
            doc.Comments.RemoveAt(1);

            //Save and launch
            string output = "RemoveAndReplaceComment.docx";
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
