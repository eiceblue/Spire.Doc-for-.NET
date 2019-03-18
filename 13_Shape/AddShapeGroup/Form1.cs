using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AddShapeGroup
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
            Section sec = doc.AddSection();

            //add a new paragraph
            Paragraph para = sec.AddParagraph();
            //add a shape group with the height and width
            ShapeGroup shapegroup = para.AppendShapeGroup(375, 462);
            shapegroup.HorizontalPosition = 180;
            //calcuate the scale ratio
            float X = (float)(shapegroup.Width / 1000.0f);
            float Y = (float)(shapegroup.Height / 1000.0f);

            Spire.Doc.Fields.TextBox txtBox = new Spire.Doc.Fields.TextBox(doc);
            txtBox.SetShapeType(ShapeType.RoundRectangle);
            txtBox.Width = 125 / X;
            txtBox.Height = 54 / Y;
            Paragraph paragraph = txtBox.Body.AddParagraph();
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            paragraph.AppendText("Start");
            txtBox.HorizontalPosition = 19/ X;
            txtBox.VerticalPosition = 27 / Y;
            txtBox.Format.LineColor = Color.Green;
            shapegroup.ChildObjects.Add(txtBox);

            ShapeObject arrowLineShape = new ShapeObject(doc, ShapeType.DownArrow);
            arrowLineShape.Width = 16 / X;
            arrowLineShape.Height = 40 / Y;
            arrowLineShape.HorizontalPosition = 69 / X;
            arrowLineShape.VerticalPosition = 87 / Y;
            arrowLineShape.StrokeColor = Color.Purple;
            shapegroup.ChildObjects.Add(arrowLineShape);

            txtBox = new Spire.Doc.Fields.TextBox(doc);
            txtBox.SetShapeType(ShapeType.Rectangle);
            txtBox.Width = 125 / X;
            txtBox.Height = 54 / Y;
            paragraph = txtBox.Body.AddParagraph();
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            paragraph.AppendText("Step 1");
            txtBox.HorizontalPosition = 19/ X;
            txtBox.VerticalPosition = 131/ Y;
            txtBox.Format.LineColor = Color.Blue;
            shapegroup.ChildObjects.Add(txtBox);

            arrowLineShape = new ShapeObject(doc, ShapeType.DownArrow);
            arrowLineShape.Width = 16 / X;
            arrowLineShape.Height = 40 / Y;
            arrowLineShape.HorizontalPosition = 69 / X;
            arrowLineShape.VerticalPosition = 192 / Y;
            arrowLineShape.StrokeColor = Color.Purple;
            shapegroup.ChildObjects.Add(arrowLineShape);

            txtBox = new Spire.Doc.Fields.TextBox(doc);
            txtBox.SetShapeType(ShapeType.Parallelogram);
            txtBox.Width = 149 / X;
            txtBox.Height = 59/ Y;
            paragraph = txtBox.Body.AddParagraph();
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            paragraph.AppendText("Step 2");
            txtBox.HorizontalPosition = 7 / X;
            txtBox.VerticalPosition = 236/ Y;
            txtBox.Format.LineColor = Color.BlueViolet;
            shapegroup.ChildObjects.Add(txtBox);

            arrowLineShape = new ShapeObject(doc, ShapeType.DownArrow);
            arrowLineShape.Width = 16 / X;
            arrowLineShape.Height = 40/ Y;
            arrowLineShape.HorizontalPosition = 66 / X;
            arrowLineShape.VerticalPosition = 300 / Y;
            arrowLineShape.StrokeColor = Color.Purple;
            shapegroup.ChildObjects.Add(arrowLineShape);

            txtBox = new Spire.Doc.Fields.TextBox(doc);
            txtBox.SetShapeType(ShapeType.Rectangle);
            txtBox.Width = 125 / X;
            txtBox.Height = 54 / Y;
            paragraph = txtBox.Body.AddParagraph();
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            paragraph.AppendText("Step 3");
            txtBox.HorizontalPosition = 19 / X;
            txtBox.VerticalPosition = 345 / Y;
            txtBox.Format.LineColor = Color.Blue;
            shapegroup.ChildObjects.Add(txtBox);



            //save the document
            doc.SaveToFile("ShapeGroup.docx", FileFormat.Docx2010);

            FileViewer("ShapeGroup.docx");
        }
        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
