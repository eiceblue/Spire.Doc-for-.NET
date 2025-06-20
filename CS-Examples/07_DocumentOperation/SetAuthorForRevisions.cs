using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetAuthorForRevisions
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object
            Document document = new Document();

            // Load the document from the specified file path
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\ExtractText.docx");

            //Get the first section
            Section section=document.Sections[0];

            // Start track revisions
            document.StartTrackRevisions("test");

            // Set author for deleted revision
            Paragraph para =document.LastParagraph;
            para.Text = "";
            for (int i = 0; i < para.ChildObjects.Count; i++)
            {
                TextRange textRange = para.ChildObjects[i] as TextRange;
                if (textRange.IsDeleteRevision)
                {
                    textRange.DeleteRevision.Author = "user1";
                }
            }

            // Set author for inserted revision
            Paragraph paragraph = section.AddParagraph();
            TextRange range = paragraph.AppendText("Added text");
            range.InsertRevision.Author = "user2";

            // Stop track revisions
            document.StopTrackRevisions();  

            // Save the file
            document.SaveToFile("SetAuthorForRevisions_out.docx", FileFormat.Docx);

            // Dispose of the Document object 
            document.Dispose();


            WordDocViewer("SetAuthorForRevisions_out.docx");

            this.Close();
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
