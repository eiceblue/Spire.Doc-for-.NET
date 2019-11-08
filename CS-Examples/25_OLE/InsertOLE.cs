using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertOLE
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {  //create a document
            Document doc = new Document();

            //add a section
            Section sec = doc.AddSection();

            //add a paragraph
            Paragraph par = sec.AddParagraph();

            //load the image
            DocPicture picture = new DocPicture(doc);
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\excel.png");
            picture.LoadImage(image);

            //insert the OLE
            DocOleObject obj = par.AppendOleObject(@"..\..\..\..\..\..\Data\example.xlsx", picture, OleObjectType.ExcelWorksheet);
            doc.SaveToFile("InsertOLE.docx", FileFormat.Docx2013);

            FileViewer("InsertOLE.docx");
        }
        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
