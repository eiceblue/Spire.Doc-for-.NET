using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertOLEAsIconViaStream
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string output = "InsertOLEAsIconViaStream.docx";

            //Create word document
            Document doc = new Document();
            //add a section
            Section sec = doc.AddSection();
            //add a paragraph
            Paragraph par = sec.AddParagraph();

            //ole stream
            Stream stream = File.OpenRead(@"..\..\..\..\..\..\Data\example.zip");

            //load the image
            DocPicture picture = new DocPicture(doc);
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\example.png");
            picture.LoadImage(image);

            //insert the OLE from stream
            DocOleObject obj = par.AppendOleObject(stream, picture, "zip");

            //display as icon
            obj.DisplayAsIcon = true;

            doc.SaveToFile(output, FileFormat.Docx2013);

            //Launching the Word file.
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
