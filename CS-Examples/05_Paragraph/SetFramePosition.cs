using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;
using System.Text;

namespace SetFramePosition
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

            //Load the document from disk
            document.LoadFromFile(@"..\..\..\..\..\..\Data\TextInFrame.docx");

            //Get a paragraph
            Paragraph paragraph = document.Sections[0].Paragraphs[0];

            //Set the Frame's position
            if (paragraph.Frame.IsFrame)
            {
                paragraph.Frame.SetHorizontalPosition(150f);
                paragraph.Frame.SetVerticalPosition(150f);
            }

            //Save to file
            String result = "SetFramePosition_result.docx";
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the file
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
