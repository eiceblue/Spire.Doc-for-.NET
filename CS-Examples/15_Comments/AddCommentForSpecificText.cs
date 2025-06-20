using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AddCommentForSpecificText
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
			Document document = new Document();

			//Load the file from disk
			document.LoadFromFile(@"..\..\..\..\..\..\Data\CommentTemplate.docx");

			//Insert comments
			InsertComments(document, "development");

			//Save the document.
			document.SaveToFile("AddCommentForTextRange.docx", FileFormat.Docx);

			//Dispose the document
			document.Dispose();

            //Launch the Word file.
            WordDocViewer("AddCommentForTextRange.docx");
        }

        private void InsertComments(Document doc,string keystring)
        {
            //Find the key string
			TextSelection find = doc.FindString(keystring, false, true);

			//Create the commentmarkStart and commentmarkEnd
			CommentMark commentmarkStart = new CommentMark(doc);

			//Set the comment Id
			commentmarkStart.CommentId = 1;

			//Set the start type
			commentmarkStart.Type = CommentMarkType.CommentStart;

			CommentMark commentmarkEnd = new CommentMark(doc);
			commentmarkEnd.CommentId = 1;
			commentmarkEnd.Type = CommentMarkType.CommentEnd;

			//Add the content for comment
			Comment comment = new Comment(doc);

			//Add the text to the paragraph
			comment.Body.AddParagraph().Text = "Test comments";

			//Add author information
			comment.Format.Author = "E-iceblue";

			//Get the textRange
			TextRange range = find.GetAsOneRange();

			//Get its paragraph
			Paragraph para = range.OwnerParagraph;

			//Get the index of textRange 
			int index = para.ChildObjects.IndexOf(range);

			//Add comment
			para.ChildObjects.Add(comment);

			//Insert the commentmarkStart and commentmarkEnd
			para.ChildObjects.Insert(index, commentmarkStart);
			para.ChildObjects.Insert(index + 2, commentmarkEnd);
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
