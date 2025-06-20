using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace StartTrackRevisions
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

            // Start track revisions
            document.StartTrackRevisions("User01", DateTime.Now);

            // Get the first paragraph and add content
            document.Sections[0].Paragraphs[0].AppendText("User01 add new Text!");

            // Delete a paragraph
            document.Sections[0].Paragraphs.RemoveAt(2);

            // Stop track revisions
            document.StopTrackRevisions();

            // Save the file
            document.SaveToFile("StartTrackRevisions_out.docx", FileFormat.Docx);

            // Dispose of the Document object 
            document.Dispose();

            WordDocViewer("StartTrackRevisions_out.docx");

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
