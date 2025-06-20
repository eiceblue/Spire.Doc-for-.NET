using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CreateImageHyperlink
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
       
			// Specify the input file path for the template document
			string input = @"..\..\..\..\..\..\Data\BlankTemplate.docx";

			// Create a new Document object
			Document doc = new Document();

			// Load the template document from the specified file path
			doc.LoadFromFile(input);

			// Get the first section of the document
			Section section = doc.Sections[0];

			// Add a new paragraph in the section
			Paragraph paragraph = section.AddParagraph();

			// Load an image from the specified file path
			Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png");

			// Create a new DocPicture object with the loaded image
			DocPicture picture = new DocPicture(doc);

			// Load the image into the DocPicture object
			picture.LoadImage(image);

			// Append a hyperlink to the paragraph with the specified URL and the picture as the display element
			paragraph.AppendHyperlink("https://www.e-iceblue.com/Introduce/word-for-net-introduce.html", picture, HyperlinkType.WebLink);

			// Specify the output file path for the generated document
			string output = "CreateImageHyperlink.docx";

			// Save the document to the output file path in DOCX format
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the document object to free up resources
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
