using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.SmartArts;

namespace ModifySmartArt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class to represent a Word document
            Document document = new Document();

            // Load an existing Word document containing SmartArt from the specified file path
            document.LoadFromFile(@"..\..\..\..\..\..\Data\SmartArt.docx");

            // Get the first section of the loaded document
            Section section = document.Sections[0];

            // Get the first paragraph from the first section, which is expected to contain the SmartArt
            Paragraph paragraph = section.Paragraphs[0];

            // Retrieve the first child object from the paragraph and cast it as a Shape (the SmartArt container)
            Spire.Doc.Fields.Shapes.Shape shape1 = paragraph.ChildObjects[0] as Spire.Doc.Fields.Shapes.Shape;

            // Access the SmartArt object from the retrieved shape
            SmartArt smartArt = shape1.SmartArt;

            // Set the background fill type of the SmartArt to a solid color
            smartArt.BackgroundFill.FillType = FillType.Solid;

            // Set the background color of the SmartArt to a light orange/peach color using ARGB values
            smartArt.BackgroundFill.Color = Color.FromArgb(255, 242, 169, 132);

            // Get the first node (main shape) of the SmartArt graphic
            SmartArtNode node = smartArt.Nodes[0];

            // Set the text content of the first node to "Goals"
            node.Text = "Goals";

            // Get the shape properties of the first node to customize its appearance
            SmartArtShapeProperties shape = node.ShapeProperties[0];

            // Set the fill type of the shape to a solid color
            shape.Fill.FillType = FillType.Solid;

            // Set the fill color of the shape to a purple color using ARGB values
            shape.Fill.Color = Color.FromArgb(255, 160, 43, 147);

            // Set the fill type of the shape's border (line format) to a solid color
            shape.LineFormat.Fill.FillType = FillType.Solid;

            // Set the border color of the shape to the same purple color
            shape.LineFormat.Fill.Color = Color.FromArgb(255, 160, 43, 147);

            // Get the first child node of the "Goals" node
            SmartArtNode childNode = node.ChildNodes[0];

            // Set the text content of the child node to a descriptive sentence
            childNode.Text = "Set clear goals to the team.";

            // Set the border fill type of the child node's shape to a solid color
            childNode.ShapeProperties[0].LineFormat.Fill.FillType = FillType.Solid;

            // Set the border color of the child node's shape to the same purple color
            childNode.ShapeProperties[0].LineFormat.Fill.Color = Color.FromArgb(255, 160, 43, 147);

            // Get the second main node of the SmartArt graphic
            node = smartArt.Nodes[1];

            // Set the text content of the second node to "Progress"
            node.Text = "Progress";

            // Get the third main node of the SmartArt graphic
            node = smartArt.Nodes[2];

            // Set the text content of the third node to "Result"
            node.Text = "Result";

            // Get the shape properties of the third node to customize its appearance
            shape = node.ShapeProperties[0];

            // Set the fill type of the third node's shape to a solid color
            shape.Fill.FillType = FillType.Solid;

            // Set the fill color of the third node's shape to a green color using ARGB values
            shape.Fill.Color = Color.FromArgb(255, 78, 167, 46);

            // Set the fill type again (redundant but present in original code)
            shape.Fill.FillType = FillType.Solid;

            // Set the border color of the third node's shape to the same green color
            shape.LineFormat.Fill.Color = Color.FromArgb(255, 78, 167, 46);

            // Set the border fill type of the first child node under the "Result" node to a solid color
            node.ChildNodes[0].ShapeProperties[0].LineFormat.Fill.FillType = FillType.Solid;

            // Set the border color of the first child node under the "Result" node to the same green color
            node.ChildNodes[0].ShapeProperties[0].LineFormat.Fill.Color = Color.FromArgb(255, 78, 167, 46);

            // Define the output file name for saving the modified document
            string result = "ModifySmartArt.docx";

            // Save the modified document to a file in Docx2016 format
            document.SaveToFile(result, FileFormat.Docx2016);

            // Close the document to release any file handles or resources
            document.Close();

            // Dispose of the document object to free up system memory
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
