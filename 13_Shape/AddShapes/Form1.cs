using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AddShapes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create Word document.
            Document doc = new Document();
            Section sec = doc.AddSection();
            Paragraph para = sec.AddParagraph();
            int x = 60, y = 40, lineCount = 0;
            for (int i = 1; i < 20; i++)
            {
                if (lineCount > 0 && lineCount % 8 == 0)
                {
                    para.AppendBreak(BreakType.PageBreak);
                    x = 60;
                    y = 40;
                    lineCount = 0;
                }
                //Add shape and set its size and position.
                ShapeObject shape = para.AppendShape(50, 50, (ShapeType)i);
                shape.HorizontalOrigin = HorizontalOrigin.Page;
                shape.HorizontalPosition = x;
                shape.VerticalOrigin = VerticalOrigin.Page;
                shape.VerticalPosition = y + 50;
                x = x + (int)shape.Width + 50;
                if (i > 0 && i % 5 == 0)
                {
                    y = y + (int)shape.Height + 120;
                    lineCount++;
                    x = 60;
                }

            }
            doc.SaveToFile("AddShape.docx", FileFormat.Docx);

            //Launch Word file.
            WordDocViewer("AddShape.docx");
        }

        private void WordDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}