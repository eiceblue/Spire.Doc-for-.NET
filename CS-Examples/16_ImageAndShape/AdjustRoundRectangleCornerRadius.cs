using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Fields.Shapes;
using Spire.Doc.Interface;

namespace AdjustRoundRectangleCornerRadius
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Load the existing Word document
            Document document = new Document(@"..\..\..\..\..\..\Data\AdjustRoundRectangleCornerRadius.docx");

            // Get the first section of the document
            Section section = document.Sections[0];

            // Iterate through all child objects in the section's body
            foreach (DocumentObject obj in section.Body.ChildObjects)
            {
                // Check if the current object is a paragraph
                if (obj is Paragraph)
                {
                    // Cast the object to a Paragraph
                    Paragraph paragraph = (Paragraph)obj;

                    // Iterate through all child objects within the paragraph
                    foreach (DocumentObject Cobj in paragraph.ChildObjects)
                    {
                        // Check if the current child object is a Shape
                        if (Cobj is Shape)
                        {
                            // Cast the child object to a ShapeObject
                            ShapeObject shape = (ShapeObject)Cobj;

                            // Check if the shape type is a Round Rectangle
                            if (shape.ShapeType == ShapeType.RoundRectangle)
                            {
                                // Get the current corner radius of the round rectangle
                                double cornerRadius = shape.AdjustHandles.GetRoundRectangleCornerRadius();

                                // Adjust the corner radius of the round rectangle to 20
                                shape.AdjustHandles.AdjustRoundRectangle(20);
                            }
                        }
                    }
                }
            }

            // Define the output file name
            string result = "AdjustRoundRectangleCornerRadius-result.docx";

            // Save the modified document to a new file in Docx 2016 format
            document.SaveToFile(result, FileFormat.Docx2016);

            // Dispose of the document object to free resources
            document.Dispose();

            WordDocViewer(result);
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
