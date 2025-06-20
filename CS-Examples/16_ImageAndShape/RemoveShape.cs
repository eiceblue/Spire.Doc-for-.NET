using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace RemoveShape
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

			//Get all the child objects of paragraph
			foreach (Paragraph para in section.Paragraphs)
			{
				for (int i = 0; i < para.ChildObjects.Count; i++)
				{
					//If the child objects is shape object
					if (para.ChildObjects[i] is ShapeObject)
					{
						//Remove the shape object
						para.ChildObjects.RemoveAt(i);
					}
				}
			}

			//Save the document
			string output = "RemoveShape.docx";
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
