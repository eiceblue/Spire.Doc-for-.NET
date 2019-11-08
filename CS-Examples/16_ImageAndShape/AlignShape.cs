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
            //Load Document
            string input = @"..\..\..\..\..\..\Data\Shapes.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            Section section = doc.Sections[0];

            foreach (Paragraph para in section.Paragraphs)
            {
                foreach (DocumentObject obj in para.ChildObjects)
                {
                    if (obj is ShapeObject)
                    {
                        //Set the horizontal alignment as center
                        (obj as ShapeObject).HorizontalAlignment = ShapeHorizontalAlignment.Center;

                        ////Set the vertical alignment as top
                        //(obj as ShapeObject).VerticalAlignment = ShapeVerticalAlignment.Top;
                    }
                }
            }

            //Save and launch document
            string output = "AlignShape.docx";
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
