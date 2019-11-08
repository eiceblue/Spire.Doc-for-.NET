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
            //Load Document
            string input = @"..\..\..\..\..\..\Data\Shapes.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

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

            //Save and launch document
            string output = "RemoveShape.docx";
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
