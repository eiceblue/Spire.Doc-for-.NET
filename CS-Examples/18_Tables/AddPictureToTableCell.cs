using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Fields;

namespace AddPictureToTableCell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\TableTemplate.docx";

			//Create a Word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the first table from the first section of the document
			Table table1 = (Table)doc.Sections[0].Tables[0];

			//Add a picture to the specified table cell
			DocPicture picture = table1.Rows[1].Cells[2].Paragraphs[0].AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png"));

			//Set picture width
			picture.Width = 100;

			//Set picture height
			picture.Height = 100;

			//Save the document
			string output = "AddPictureToTableCell.docx";
			doc.SaveToFile(output, FileFormat.Docx);

			//Dispose the document
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
