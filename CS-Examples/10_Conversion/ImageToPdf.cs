using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;
using Spire.Doc.Fields;
namespace ImageToPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\Image.png";
			//Create a new document
			Document doc = new Document();

			//Create a new section
			Section section = doc.AddSection();

			//Create a new paragraph
			Paragraph paragraph = section.AddParagraph();

			//Add a picture for paragraph
			paragraph.AppendPicture(input);

			//Set A4 page size
			section.PageSetup.PageSize = PageSize.A4;

			//Set the page margins
			section.PageSetup.Margins.Top = 10f;
			section.PageSetup.Margins.Left = 25f;

			string result = "ImageToPdf.pdf";
			
			//Save to file
			doc.SaveToFile(result, FileFormat.PDF);

			//Dispose the Document
			doc.Dispose();
			
			//Launch the document
            Viewer(result);
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
