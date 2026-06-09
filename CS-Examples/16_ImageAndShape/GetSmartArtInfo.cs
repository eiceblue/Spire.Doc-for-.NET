using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.SmartArts;

namespace GetSmartArtInfo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            // Create a StringBuilder instance to efficiently accumulate the extracted SmartArt information
            StringBuilder builder = new StringBuilder();

            // Open the Word document containing SmartArt using a 'using' statement for automatic resource disposal
            using (Document document = new Document(@"..\..\..\..\..\..\Data\SmartArt.docx"))
            {
                // Iterate through every section in the document
                foreach (Section section in document.Sections)
                {
                    // Skip the section if it or its paragraph collection is null to avoid exceptions
                    if (section?.Paragraphs == null) continue;

                    // Iterate through every paragraph in the current section
                    foreach (Paragraph paragraph in section.Paragraphs)
                    {
                        // Iterate through all child objects (elements) contained within the paragraph
                        foreach (var childObj in paragraph.ChildObjects)
                        {
                            // Check if the object is a Shape and if it specifically contains a SmartArt graphic
                            if (childObj is Spire.Doc.Fields.Shapes.Shape shape && shape.HasSmartArt)
                            {
                                // Retrieve the SmartArt object from the shape
                                SmartArt smartArt = shape.SmartArt;

                                // Skip processing if the SmartArt object is unexpectedly null
                                if (smartArt == null) continue;

                                // Append the type of the SmartArt graphic to the result string
                                builder.AppendLine($"SmartArtType£º{smartArt.SmartArtType}");

                                // Call a helper method to extract and append background formatting details
                                ExtractSmartArtBackgroundInfo(smartArt, builder);

                                // Call a recursive helper method to traverse nodes and extract their text and properties
                                TraverseSmartArtNodes(smartArt.Nodes, builder, 0);

                                // Append a separator line to distinguish this SmartArt block from the next one
                                builder.AppendLine("----------------------------------------rn");
                            }
                        }
                    }
                }
            }

            // Define the file path for the output text file
            string result = "GetSmartArtInfo.txt";

            // Write the accumulated text content from the StringBuilder to the file
            File.WriteAllText(result, builder.ToString());
        }
        // Define a method to extract and append SmartArt background fill information to the StringBuilder
        public void ExtractSmartArtBackgroundInfo(SmartArt smartArt, StringBuilder builder)
        {
            // Return immediately if the background fill type is set to NoFill (transparent/none)
            if (smartArt?.BackgroundFill.FillType == FillType.NoFill)
            {
                return;
            }

            // Convert the background fill type enum to its string representation
            string bgFillType = smartArt.BackgroundFill.FillType.ToString();

            // Check if the background color is empty; if so, use a placeholder text, otherwise convert color to string
            string bgColor = smartArt.BackgroundFill.Color == Color.Empty
                ? "No color"
                : smartArt.BackgroundFill.Color.ToString();

            // Append the background fill type and color information to the StringBuilder with newlines
            builder.AppendLine($"BackgroundFill_filltype£º{bgFillType}\nBackgroundFill_color£º{bgColor}");
        }

        // Define a recursive method to traverse SmartArt nodes and extract their text and properties
        public static void TraverseSmartArtNodes(SmartArtNodeCollection nodes, StringBuilder builder, int level)
        {
            // Exit the method if the node collection is null or contains no nodes
            if (nodes == null || nodes.Count == 0) return;

            // Iterate through each node in the current collection
            for (int nodeIdx = 0; nodeIdx < nodes.Count; nodeIdx++)
            {
                // Get the current node from the collection
                SmartArtNode node = nodes[nodeIdx];

                // Skip to the next iteration if the current node is null
                if (node == null) continue;

                // Trim whitespace from the node text, or use a placeholder if the text is null
                string nodeText = node.Text != null ? node.Text.Trim() : "Empty Text";

                // Skip this node if the text is just a carriage return or effectively empty
                if (nodeText == "\r" || string.IsNullOrEmpty(nodeText)) continue;

                // Declare a variable to hold the prefix string based on the hierarchy level
                string nodePrefix;

                // Determine the appropriate prefix string based on the current depth level in the hierarchy
                switch (level)
                {
                    case 0:
                        nodePrefix = "smartArt.Nodes";
                        break;
                    case 1:
                        nodePrefix = "smartArt.Nodes.ChildNodes";
                        break;
                    case 2:
                        nodePrefix = "smartArt.Nodes.ChildNodes.ChildNodes";
                        break;
                    default:
                        nodePrefix = $"smartArt.Nodes.Level{level}";
                        break;
                }

                // Append the formatted node index and text content to the StringBuilder
                builder.AppendLine($"{nodePrefix}_{nodeIdx}£º{nodeText}");

                // Call a helper method to extract and append specific properties of the current node
                ExtractSmartArtNodeProperties(node, builder);

                // Recursively call this method to process any child nodes, incrementing the level counter
                TraverseSmartArtNodes(node.ChildNodes, builder, level + 1);
            }
        }
        // Define a method to extract and append shape formatting properties (fill and border) of a SmartArt node
        public static void ExtractSmartArtNodeProperties(SmartArtNode node, StringBuilder builder)
        {
            // First, call the helper method to extract font and text formatting properties
            ExtractFontProperties(node, builder);

            // Exit the method if the node has no shape properties or the collection is empty
            if (node.ShapeProperties == null || node.ShapeProperties.Count == 0)
            {
                return;
            }

            // Retrieve the first shape property object from the node to access its visual settings
            SmartArtShapeProperties shapeProps = node.ShapeProperties[0];

            // Check if the shape has a fill type other than "NoFill" (i.e., it has a solid color or picture)
            if (shapeProps?.Fill.FillType != FillType.NoFill)
            {
                // Convert the fill type enum to its string representation
                string nodeFillType = shapeProps.Fill.FillType.ToString();

                // Determine the color string: use a specific message for pictures, otherwise get the color value
                string nodeColor = nodeFillType == "Picture"
                    ? "(Picture type, without color acquisition)"
                    : shapeProps.Fill.Color.ToString();

                // Append the fill type and color information to the StringBuilder with indentation
                builder.AppendLine($"\tfilltype£º{nodeFillType}\n\tcolor£º{nodeColor}");
            }

            // Check if the shape's border (line format) has a fill type other than "NoFill" and a valid color
            if (shapeProps.LineFormat?.Fill.FillType != FillType.NoFill && shapeProps.LineFormat.Fill.Color != Color.Empty)
            {
                // Convert the border fill type enum to its string representation
                string lineFillType = shapeProps.LineFormat.Fill.FillType.ToString();

                // Convert the border color to its string representation
                string lineColor = shapeProps.LineFormat.Fill.Color.ToString();

                // Append the border fill type and color information to the StringBuilder with indentation
                builder.AppendLine($"\tline_filltype£º{lineFillType}\n\tline_color£º{lineColor}");
            }
        }

        // Define a private helper method to extract and append font formatting properties of the node's text
        private static void ExtractFontProperties(SmartArtNode node, StringBuilder builder)
        {
            // Return immediately if the node, its paragraphs collection, or the paragraphs themselves are null or empty
            if (node?.Paragraphs == null || node.Paragraphs.Count == 0)
                return;

            // Get the first paragraph from the node, which typically contains the main text
            var paragraph = node.Paragraphs[0];

            // Return if the paragraph's child objects collection is null or empty
            if (paragraph?.ChildObjects == null || paragraph.ChildObjects.Count == 0)
                return;

            // Attempt to cast the first child object as a TextRange to access character formatting
            var textRange = paragraph.ChildObjects[0] as Spire.Doc.Fields.TextRange;

            // Return if the cast fails (i.e., the object is not a TextRange)
            if (textRange == null)
                return;

            // Retrieve the CharacterFormat object which holds all font-related settings
            var charFormat = textRange.CharacterFormat;

            // Extract the font name from the character format
            string fontName = charFormat.FontName;

            // Extract the font size from the character format
            float fontSize = charFormat.FontSize;

            // Extract the text color from the character format
            Color fontColor = charFormat.TextColor;

            // Extract the font style (bold, italic, etc.) from the character format
            Spire.Doc.Publics.Drawing.FontStyle fontstyle = charFormat.FontStyle;

            // Initialize a flag to track whether any valid font information was found
            bool hasValidFontInfo = false;

            // Create a new StringBuilder to accumulate the font property strings efficiently
            StringBuilder fontInfoBuilder = new StringBuilder();

            // Check if the font name is not empty and append it if valid
            if (!string.IsNullOrEmpty(fontName))
            {
                fontInfoBuilder.Append($"\tfont_name£º{fontName}");
                hasValidFontInfo = true;
            }

            // Check if the font size is greater than zero and append it if valid
            if (fontSize > 0)
            {
                if (hasValidFontInfo)
                    fontInfoBuilder.Append($"\tfont_size£º{fontSize}pt");
                hasValidFontInfo = true;
            }

            // Check if the font color is not empty and append it if valid
            if (fontColor != Color.Empty)
            {
                if (hasValidFontInfo)
                    fontInfoBuilder.Append($"\tfont_color£º{fontColor}");
                hasValidFontInfo = true;
            }

            // Append the font style information regardless of other fields
            fontInfoBuilder.Append($"\tfont_style£º{fontstyle}");

            // If any valid font information was collected, append the entire string to the main builder
            if (hasValidFontInfo)
            {
                builder.AppendLine(fontInfoBuilder.ToString());
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
