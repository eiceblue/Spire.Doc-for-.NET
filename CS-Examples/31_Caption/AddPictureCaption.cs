using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AddPictureCaption
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of Document
            Document document = new Document();

            // Add a new section to the document
            Section section = document.AddSection();

            // Add a paragraph to the section
            Paragraph par1 = section.AddParagraph();
            par1.Format.AfterSpacing = 10;

            // Append an image (picture) to the paragraph from the specified file path
            DocPicture pic1 = par1.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png"));
            pic1.Height = 100;
            pic1.Width = 120;

            // Set the caption numbering format to "Number" and add a caption below the picture
            CaptionNumberingFormat format = CaptionNumberingFormat.Number;
            pic1.AddCaption("Figure", format, CaptionPosition.BelowItem);

            // Add another paragraph to the section
            Paragraph par2 = section.AddParagraph();

            // Append another image (picture) to the paragraph from the specified file path
            DocPicture pic2 = par2.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Word.png"));
            pic2.Height = 100;
            pic2.Width = 120;

            // Add a caption below the second picture
            pic2.AddCaption("Figure", format, CaptionPosition.BelowItem);

            // Enable field updating in the document
            document.IsUpdateFields = true;

            // Specify the output file name and format (Docx)
            string output = "AddPictureCaption_result.docx";
            document.SaveToFile(output, FileFormat.Docx);

            // Dispose of the document object when finished using it
            document.Dispose();
			
            //Launching the file
            WordDocViewer(output);

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
