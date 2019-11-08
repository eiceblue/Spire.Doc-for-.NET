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
            //Load Document
            string input = @"..\..\..\..\..\..\Data\Shapes.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the first section and the first paragraph that contains the shape
            Section section = doc.Sections[0];
            Paragraph para = section.Paragraphs[0];

            //Get the second shape and reset the width and height for the shape
            ShapeObject shape = para.ChildObjects[1] as ShapeObject;
            shape.Width = 200;
            shape.Height = 200; 

            //Save and launch document
            string output = "ResetShapeSize.docx";
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
