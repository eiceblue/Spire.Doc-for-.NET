using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.SmartArts;

namespace AddSmartArt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private const int SmartArtDefaultWidth = 432;
        private const int SmartArtDefaultHeight = 252;
        private const float TitleFontSize = 28f;
        private const float DefaultNodeFontSize = 15f;
        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class to represent a Word document
            Document document = new Document();

            // Add a new section to the document, which serves as a container for content like paragraphs and SmartArt
            Section section = document.AddSection();

            // Initialize a list to store various types of SmartArt graphics that will be added to the document
            List<SmartArtType> smartArtTypes = new List<SmartArtType>();

            // Add a vertical chevron list SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.VerticalChevronList);

            // Add a square accent list SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.SquareAccentList);

            // Add an alternating hexagons SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.AlternatingHexagons);

            // Add a horizontal bullet list SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.HorizontalBulletList);

            // Add a segmented process SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.SegmentedProcess);

            // Add a vertical bending process SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.VerticalBendingProcess);

            // Add a step down process SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.StepDownProcess);

            // Add a circle accent timeline SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.CircleAccentTimeLine);

            // Add a block cycle SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.BlockCycle);

            // Add a segmented cycle SmartArt type to the collection
            smartArtTypes.Add(SmartArtType.SegmentedCycle);

            // Iterate through each SmartArt type in the list to create and configure SmartArt graphics
            foreach (SmartArtType smartArtType in smartArtTypes)
            {
                // Call a helper method to create a title paragraph with the name of the current SmartArt type
                CreateTitleParagraph(section, smartArtType.ToString());

                // Add a new paragraph to the section which will contain the SmartArt graphic
                var paragraph = section.AddParagraph();

                // Call a helper method to create and insert the SmartArt graphic into the paragraph
                var smartArt = CreateSmartArt(paragraph, smartArtType);

                // Get the shape properties of the first node's first shape to customize its appearance
                SmartArtShapeProperties shapeSmartArt = smartArt.Nodes[0].ShapeProperties[0];

                // Set the fill type of the shape to a solid color
                shapeSmartArt.Fill.FillType = FillType.Solid;

                // Set the fill color of the shape to orange using ARGB values (255, 165, 0)
                shapeSmartArt.Fill.Color = Color.FromArgb(255, 255, 165, 0);

                // Set the text and font size for the first node of the SmartArt graphic
                SetSmartArtNodeText(smartArt.Nodes[0], "TextTest_1", 15f);

                // Add a child node to the first node with specified font size and text
                AddSmartArtChildNode(smartArt.Nodes[0], 15f, "ChildNodeTest_1.");

                // Set the text and a larger font size for the second node of the SmartArt graphic
                SetSmartArtNodeText(smartArt.Nodes[1], "TextTest_2", 25f);

                // Add a child node to the second node with specified font size and text
                AddSmartArtChildNode(smartArt.Nodes[1], 15f, "ChildNodeTest_2.");
            }

            // Define the file path and name for saving the generated Word document
            String result = "AddSmartArt.docx";

            // Save the document to a file in Docx2016 format
            document.SaveToFile(result, FileFormat.Docx2016);

            // Close the document to release any file handles or resources
            document.Close();

            // Dispose of the document object to free up system memory
            document.Dispose();

            WordDocViewer(result);
        }
        // Define a method to create and return a title paragraph with specified text and formatting
        private Spire.Doc.Documents.Paragraph CreateTitleParagraph(Section section, string titleText, float fontSize = TitleFontSize)
        {
            // Add a new paragraph to the given section
            var paragraph = section.AddParagraph();

            // Set the horizontal alignment of the paragraph to center
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            // Append the title text to the paragraph and get the TextRange object for formatting
            var textRange = paragraph.AppendText(titleText);

            // Set the font size of the title text
            textRange.CharacterFormat.FontSize = fontSize;

            // Set the font name of the title text to Times New Roman
            textRange.CharacterFormat.FontName = "Times New Roman";

            // Add two empty paragraphs after the title to create vertical spacing
            section.AddParagraph();
            section.AddParagraph();

            // Return the created title paragraph
            return paragraph;
        }

        // Define a method to create and insert a SmartArt graphic into a paragraph
        private SmartArt CreateSmartArt(Spire.Doc.Documents.Paragraph paragraph, SmartArtType smartArtType, int width = SmartArtDefaultWidth, int height = SmartArtDefaultHeight)
        {
            // Set the horizontal alignment of the paragraph containing the SmartArt to center
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            // Append a SmartArt shape of the specified type and dimensions to the paragraph
            Spire.Doc.Fields.Shapes.Shape shape = paragraph.AppendSmartArt(smartArtType, width, height);

            // Return the SmartArt object from the created shape
            return shape.SmartArt;
        }

        // Define a method to set the text and font size for a SmartArt node
        private void SetSmartArtNodeText(SmartArtNode node, string text, float fontSize = DefaultNodeFontSize)
        {
            // Exit the method early if the node is null or the text is empty
            if (node == null || string.IsNullOrEmpty(text)) return;

            // Set the text content of the SmartArt node
            node.Text = text;

            // Check if the node has paragraphs and the first paragraph has child objects (like TextRange)
            if (node.Paragraphs.Count > 0 && node.Paragraphs[0].ChildObjects.Count > 0)
            {
                // Cast the first child object to TextRange and set its font size
                ((Spire.Doc.Fields.TextRange)node.Paragraphs[0].ChildObjects[0]).CharacterFormat.FontSize = fontSize;
            }
        }

        // Define a method to add child nodes to a parent SmartArt node and set their text
        private void AddSmartArtChildNode(SmartArtNode parentNode, float fontSize, params string[] childTexts)
        {
            // Exit the method early if the parent node is null, or the child texts array is null or empty
            if (parentNode == null || childTexts == null || childTexts.Length == 0) return;

            // Ensure the parent node has enough child nodes to accommodate all the provided text strings
            while (parentNode.ChildNodes.Count < childTexts.Length)
            {
                // Add a new child node until the count matches the number of text strings
                parentNode.ChildNodes.Add();
            }

            // Iterate through the child text strings and corresponding child nodes
            for (int i = 0; i < childTexts.Length && i < parentNode.ChildNodes.Count; i++)
            {
                // Call SetSmartArtNodeText to set the text and font size for each child node
                SetSmartArtNodeText(parentNode.ChildNodes[i], childTexts[i], fontSize);
            }
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
