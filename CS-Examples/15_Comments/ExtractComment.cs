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
            //Load the document
            string input = @"..\..\..\..\..\..\Data\CommentSample.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            StringBuilder SB = new StringBuilder();

            //Traverse all comments
            foreach (Comment comment in doc.Comments)
            {
                foreach (Paragraph p in comment.Body.Paragraphs)
                {
                    SB.AppendLine(p.Text);
                }
            }

            //Save to TXT File and launch it
            string output = "ExtractComment.txt";
            File.WriteAllText(output, SB.ToString());
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
