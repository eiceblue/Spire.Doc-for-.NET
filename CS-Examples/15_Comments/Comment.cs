using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace Comment
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load the document from disk.
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\Data\CommentTemplate.docx");

            InsertComments(document.Sections[0]);

            //Save the document.
            document.SaveToFile("Output.docx",FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer("Output.docx");
        }

        private void InsertComments(Section section)
        {          
            //Insert comment.
            Paragraph paragraph = section.Paragraphs[1];
            Spire.Doc.Fields.Comment comment = paragraph.AppendComment("Spire.Doc for .NET");
            comment.Format.Author = "E-iceblue";
            comment.Format.Initial = "CM";
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
