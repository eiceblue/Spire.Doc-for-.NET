using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace FromCommentRange
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a document
            Document sourceDoc = new Document();

            //Load the document from disk.
            sourceDoc.LoadFromFile(@"..\..\..\..\..\..\Data\Comments.docx");

            //Create a destination document
            Document destinationDoc = new Document();

            //Add section for destination document
            Section destinationSec = destinationDoc.AddSection();

            //Get the first comment
            Comment comment = sourceDoc.Comments[0];
            
            //Get the paragraph of obtained comment
            Paragraph para = comment.OwnerParagraph;

            //Get index of the CommentMarkStart 
            int startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart);

            //Get index of the CommentMarkEnd
            int endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd);

            //Traverse paragraph ChildObjects
            for (int i = startIndex; i <= endIndex; i++)
            {
                //Clone the ChildObjects of source document
                DocumentObject doobj = para.ChildObjects[i].Clone();

                //Add to destination document 
                destinationSec.AddParagraph().ChildObjects.Add(doobj);
            }
            //Save the destination document
            destinationDoc.SaveToFile("Output.docx", FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer("Output.docx");
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
