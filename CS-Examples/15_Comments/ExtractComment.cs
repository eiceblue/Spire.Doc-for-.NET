using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ExtractComment
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

			//Create a StringBuilder instance
			StringBuilder SB = new StringBuilder();

			//Traverse all comments
			foreach (Comment comment in doc.Comments)
			{
				foreach (Paragraph p in comment.Body.Paragraphs)
				{
					//Append the comments to the StringBuilder instance
					SB.AppendLine(p.Text);
				}
			}

			//Save to TXT File and launch it
			string output = "ExtractComment.txt";
			File.WriteAllText(output, SB.ToString());

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
