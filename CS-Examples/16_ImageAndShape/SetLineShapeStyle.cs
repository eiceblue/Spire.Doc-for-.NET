using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetLineShapeStyle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //create a document
			Document doc = new Document();

			//Add a section
			Section sec = doc.AddSection();

			//Add a new paragraph
			Paragraph para = sec.AddParagraph();

			//Add a line shape
			ShapeObject shape = para.AppendShape(100, 100, ShapeType.Line);

			//Set style of Line shape
			shape.FillColor = Color.Orange;
			shape.StrokeColor = Color.Black;
			shape.LineStyle = ShapeLineStyle.Single;
			shape.LineDashing = LineDashing.LongDashDotDotGEL;

			//Save the document
			string output = "SetLineShapeStyle.docx";
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
