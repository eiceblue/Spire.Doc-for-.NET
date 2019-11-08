using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertPictureIntoComment
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
            string input = @"..\..\..\..\..\..\Data\CommentTemplate.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the first paragraph and insert comment
            Paragraph paragraph = doc.Sections[0].Paragraphs[2];
            Comment comment = paragraph.AppendComment("This is a comment.");
            comment.Format.Author = "E-iceblue";

            //Load a picture
            DocPicture docPicture = new DocPicture(doc);
            Image img = Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png");
            docPicture.LoadImage(img);

            //Insert the picture into the comment body
            comment.Body.AddParagraph().ChildObjects.Add(docPicture);

            //Save and launch
            string output = "InsertPictureIntoComment.docx";
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
