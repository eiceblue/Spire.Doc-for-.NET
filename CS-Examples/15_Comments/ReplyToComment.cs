using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Fields;

namespace ReplyToComment
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a word document
			Document doc = new Document();

			//Load the document from disk.
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\Comment.docx");

			//get the first comment.
			Comment comment1 = doc.Comments[0];

			//create a new comment
			Comment replyComment1 = new Comment(doc);

			//Set the author
			replyComment1.Format.Author = "E-iceblue";

			//Append text
			replyComment1.Body.AddParagraph().AppendText("Spire.Doc is a professional Word .NET library on operating Word documents.");

			//add the new comment as a reply to the selected comment.
			comment1.ReplyToComment(replyComment1);

			//Create a DocPicture instance
			DocPicture docPicture = new DocPicture(doc);

			//Load an image
			docPicture.LoadImage(Image.FromFile(@"..\..\..\..\..\..\Data\logo.png"));

			//insert a picture in the comment
			replyComment1.Body.Paragraphs[0].ChildObjects.Add(docPicture);

			//Save the document.
			doc.SaveToFile("ReplyToComment.docx", FileFormat.Docx);

			//Dispose the document
			doc.Dispose();
			
            //Launch the Word file.
            WordDocViewer("ReplyToComment.docx");
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
