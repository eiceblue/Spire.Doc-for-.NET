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
            //Create word document
            Document document = new Document();

            //Create a new section
            Section section = document.AddSection();

            //Add the first picture
            Paragraph par1 = section.AddParagraph();
            par1.Format.AfterSpacing = 10;
            DocPicture pic1 = par1.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png"));
            pic1.Height = 100;
            pic1.Width = 120;
            //Add caption to the picture
            CaptionNumberingFormat format = CaptionNumberingFormat.Number;
            pic1.AddCaption("Figure", format, CaptionPosition.BelowItem);

            //Add the second picture
            Paragraph par2 = section.AddParagraph();
            DocPicture pic2 = par2.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Word.png"));
            pic2.Height = 100;
            pic2.Width = 120;
            //Add caption to the picture
            pic2.AddCaption("Figure", format, CaptionPosition.BelowItem);

            //Update fields
            document.IsUpdateFields = true;

            //Save the file
            string output = "AddPictureCaption_result.docx";
            document.SaveToFile(output,FileFormat.Docx);

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
