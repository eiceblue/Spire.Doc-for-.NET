using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddHorizontalLine
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a Word document.
			Document doc = new Document();

			//Add a section
			Section sec = doc.AddSection();

			//Add a paragraph
			Paragraph para = sec.AddParagraph();

			//Append an horizonal line
			para.AppendHorizonalLine();

			string result = "AddHorizontalLine_result.docx";

			//Save the document
			doc.SaveToFile(result, FileFormat.Docx);

			//Dispose the document
			doc.Dispose();
            
            //Launching the MS Word file.
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
