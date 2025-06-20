using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AlignShape
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

			//Loop through the paragraphs in the section
			foreach (Paragraph para in section.Paragraphs)
			{
				// Loop through the child objects in the paragraph
				foreach (DocumentObject obj in para.ChildObjects)
				{
					if (obj is ShapeObject)
					{
						//Set the horizontal alignment as center
						(obj as ShapeObject).HorizontalAlignment = ShapeHorizontalAlignment.Center;
					}
				}
			}

			//Save the document
			string output = "AlignShape.docx";
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
