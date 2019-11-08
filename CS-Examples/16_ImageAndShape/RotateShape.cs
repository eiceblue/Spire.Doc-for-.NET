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
            //Load Document
            string input = @"..\..\..\..\..\..\Data\Shapes.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the first section
            Section section = doc.Sections[0];

            //Traverse the word document and set the shape rotation as 20
            foreach (Paragraph para in section.Paragraphs)
            {
                foreach (DocumentObject obj in para.ChildObjects)
                {
                    if (obj is ShapeObject)
                    {
                        (obj as ShapeObject).Rotation = 20.0;
                    }
                }
            }

            //Save and launch document
            string output = "RotateShape.docx";
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
