using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddSVG
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Specify the input SVG file path
            string inputSvg = "../../../../../../Data/charthtml.svg";

            // Specify the output Word document file path
            string outputFile = "addSVG.docx";

            // Create a new Document object
            Document document = new Document();

            // Add a new Section to the document
            Section section = document.AddSection();

            // Add a new Paragraph to the section
            Paragraph paragraph = section.AddParagraph();

            // Append the picture (SVG) to the paragraph
            paragraph.AppendPicture(inputSvg);

            // Save the document to the specified output file
            document.SaveToFile(outputFile, FileFormat.Docx2013);
            
            // Close the document
            document.Close();

            // Launch Word file
            WordDocViewer(outputFile);
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