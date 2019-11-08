using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ReplaceWithImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load Document
            string input = @"..\..\..\..\..\..\Data\Template.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Load the image
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png");

            //Find the string "E-iceblue" in the document
            TextSelection[] selections = doc.FindAllString("E-iceblue", true, true);
            int index = 0;
            TextRange range = null;

            //Remove the text and replace it with Image
            foreach (TextSelection selection in selections)
            {
                DocPicture pic = new DocPicture(doc);
                pic.LoadImage(image);

                range = selection.GetAsOneRange();
                index = range.OwnerParagraph.ChildObjects.IndexOf(range);
                range.OwnerParagraph.ChildObjects.Insert(index, pic);
                range.OwnerParagraph.ChildObjects.Remove(range);
            }

            //Save and launch document
            string output = "ReplaceWithImage.docx";
            doc.SaveToFile(output, FileFormat.Docx);
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
