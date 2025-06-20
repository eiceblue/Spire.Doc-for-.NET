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
            // Create a new Document object to represent the source document.
            Document sourceDoc = new Document();

            // Load the Word document from the specified file path.
            sourceDoc.LoadFromFile(@"..\..\..\..\..\..\Data\Comments.docx");

            // Create a new Document object to represent the destination document.
            Document destinationDoc = new Document();

            // Add a section to the destination document.
            Section destinationSec = destinationDoc.AddSection();

            // Get the first comment from the source document.
            Comment comment = sourceDoc.Comments[0];

            // Get the paragraph that owns the comment.
            Paragraph para = comment.OwnerParagraph;

            // Find the index of the CommentMarkStart and CommentMarkEnd within the paragraph's ChildObjects.
            int startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart);
            int endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd);

            // Iterate over the ChildObjects in the paragraph between the start and end indices.
            for (int i = startIndex; i <= endIndex; i++)
            {
                // Clone the DocumentObject at the current index.
                DocumentObject doobj = para.ChildObjects[i].Clone();

                // Add the cloned DocumentObject to a new paragraph in the destination section.
                destinationSec.AddParagraph().ChildObjects.Add(doobj);
            }

            // Save the destination document to a file.
            destinationDoc.SaveToFile("Output.docx", FileFormat.Docx);

            // Dispose of the source and destination documents to free up resources.
            sourceDoc.Dispose();
            destinationDoc.Dispose();

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
