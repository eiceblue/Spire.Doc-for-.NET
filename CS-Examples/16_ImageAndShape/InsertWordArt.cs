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

namespace InsertWordArt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {          
            //Create a Word document.
			Document doc = new Document();

			//Load Word document.
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\InsertWordArt.docx");

			//Add a paragraph.
			Paragraph paragraph = doc.Sections[0].AddParagraph();

			//Add a shape.
			ShapeObject shape = paragraph.AppendShape(250, 70, ShapeType.TextWave4);

			//Set the position of the shape.
			shape.VerticalPosition = 20;
			shape.HorizontalPosition = 80;

			//set the text of WordArt.
			shape.WordArt.Text = "Thanks for reading.";

			//Set the fill color.
			shape.FillColor = Color.Red;

			//Set the border color of the text.
			shape.StrokeColor = Color.Yellow;

			//Save docx file.
			doc.SaveToFile("WordArt.docx", FileFormat.Docx2013);

			// Dispose the document
			doc.Dispose();

            //Launch the Word file.
            FileViewer("WordArt.docx");
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
