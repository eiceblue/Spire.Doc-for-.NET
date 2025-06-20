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
            string input = @"..\..\..\..\..\..\Data\CommentSample.docx";

			//Create a word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Replace the content of the first comment
			doc.Comments[0].Body.Paragraphs[0].Replace("This is the title", "This comment is changed.", false, false);

			//Remove the second comment
			doc.Comments.RemoveAt(1);
			
			string output = "RemoveAndReplaceComment.docx";
			
			//Save the document
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
