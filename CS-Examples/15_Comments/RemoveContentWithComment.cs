using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Collections.Generic;

namespace RemoveContentWithComment
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
            Document document = new Document();

            //Load the document from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Comments.docx");

            //Get the first comment
            Comment comment = document.Comments[0];

            //Get the paragraph of obtained comment
            Paragraph para = comment.OwnerParagraph;

            //Get index of the CommentMarkStart 
            int startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart);

            //Get index of the CommentMarkEnd
            int endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd);

            //Create a list
            List<TextRange> list = new List<TextRange>();

            //Get TextRanges between the indexes
            for (int i = startIndex; i < endIndex; i++)
            {
                if (para.ChildObjects[i] is TextRange)
                {
                    list.Add(para.ChildObjects[i] as TextRange);
                }
            }

            //Insert a new TextRange
            TextRange textRange = new TextRange(document);

            //Set text is null
            textRange.Text = null;

            //Insert the new textRange
            para.ChildObjects.Insert(endIndex, textRange);

            //Remove previous TextRanges
            for (int i = 0; i < list.Count; i++)
            {
                para.ChildObjects.Remove(list[i]);
            }

            String result = "Output.docx";
            //Save the document.
            document.SaveToFile(result, FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer(result);
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
