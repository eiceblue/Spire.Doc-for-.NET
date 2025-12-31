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
        {  

			// Create a new document object
			Document doc = new Document();

			// Add a section to the document
			Section sec = doc.AddSection();

			// Add a paragraph to the section
			Paragraph par = sec.AddParagraph();

			// Create a DocPicture object and load an image from file
			DocPicture picture = new DocPicture(doc);
			Image image = Image.FromFile(@"..\..\..\..\..\..\Data\excel.png");
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            picture.LoadImage(@"..\..\..\..\..\..\Data\excel.png");
            */
            picture.LoadImage(image);

			// Append an OLE object to the paragraph with the specified file, picture, and object type (Excel worksheet)
			DocOleObject obj = par.AppendOleObject(@"..\..\..\..\..\..\Data\example.xlsx", picture, OleObjectType.ExcelWorksheet);

			// Save the document to a file in Docx2013 format
			doc.SaveToFile("InsertOLE.docx", FileFormat.Docx2013);

			// Dispose the document object
			doc.Dispose();

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
