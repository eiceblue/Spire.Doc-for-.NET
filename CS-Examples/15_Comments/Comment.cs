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
            //Create a word document
			Document document = new Document();

			//Load the file from disk
			document.LoadFromFile(@"..\..\..\..\..\..\Data\CommentTemplate.docx");

			//Insert comments
			InsertComments(document.Sections[0]);

			//Save the document.
			document.SaveToFile("Output.docx", FileFormat.Docx);

			//Dispose the document
			document.Dispose();

            //Launch the Word file.
            WordDocViewer("Output.docx");
        }

        private void InsertComments(Section section)
        {          
            //Get the second paragraph
			Paragraph paragraph = section.Paragraphs[1];

			//Add comment
			Spire.Doc.Fields.Comment comment = paragraph.AppendComment("Spire.Doc for .NET");

			//Add author information
			comment.Format.Author = "E-iceblue";

			//Set the user initials.
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
