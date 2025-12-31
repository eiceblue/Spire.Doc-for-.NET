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
            string input = @"..\..\..\..\..\..\Data\CommentTemplate.docx";

			//Create a word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the third paragraph in the first section
			Paragraph paragraph = doc.Sections[0].Paragraphs[2];

			//Add comment
			Comment comment = paragraph.AppendComment("This is a comment.");

			//Add author information
			comment.Format.Author = "E-iceblue";

			//Create a DocPicture instance
			DocPicture docPicture = new DocPicture(doc);

			//Load a picture
			docPicture.LoadImage(Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png"));
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
             docPicture.LoadImage(@"..\..\..\..\..\..\Data\E-iceblue.png");
            */

            //Insert the picture into the comment body
            comment.Body.AddParagraph().ChildObjects.Add(docPicture);

			//Save the document
			string output = "InsertPictureIntoComment.docx";
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
