using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace InsertImageIntoTextBox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a new document
            Document doc = new Document();

            Section section = doc.AddSection();
            Paragraph paragraph = section.AddParagraph();

            //Append a textbox to paragraph
            Spire.Doc.Fields.TextBox tb = paragraph.AppendTextBox(220, 220);

            //Set the position of the textbox
            tb.Format.HorizontalOrigin = HorizontalOrigin.Page;
            tb.Format.HorizontalPosition = 50;
            tb.Format.VerticalOrigin = VerticalOrigin.Page;
            tb.Format.VerticalPosition = 50;

            //Set the fill effect of textbox as picture
            tb.Format.FillEfects.Type = BackgroundType.Picture;

            //Fill the textbox with a picture
            tb.Format.FillEfects.Picture = Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png");

            //Save and launch document
            string output = "InsertImageIntoTextBox.docx";
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
