using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ResetShapeSize
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\Shapes.docx");

			//Get the first section 
			Section section = doc.Sections[0];

			//Get the first paragraph
			Paragraph para = section.Paragraphs[0];

			//Get the second shape
			ShapeObject shape = para.ChildObjects[1] as ShapeObject;

			//Reset the width and height of the shape
			shape.Width = 200;
			shape.Height = 200;

			//Save and launch document
			string output = "ResetShapeSize.docx";
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the document
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
