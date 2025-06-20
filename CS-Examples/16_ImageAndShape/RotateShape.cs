using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace RotateShape
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\Shapes.docx";

			//Create a word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the first section
			Section section = doc.Sections[0];

			//Traverse the word document 
			foreach (Paragraph para in section.Paragraphs)
			{
				foreach (DocumentObject obj in para.ChildObjects)
				{
					if (obj is ShapeObject)
					{
						//Set the shape rotation as 20
						(obj as ShapeObject).Rotation = 20.0;
					}
				}
			}

			//Save the document
			string output = "RotateShape.docx";
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
