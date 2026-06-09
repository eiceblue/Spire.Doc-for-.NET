using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemoveSmartArt
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

            // Load an existing Word document from the specified file path
            document.LoadFromFile(@"..\..\..\..\..\..\Data\SmartArt.docx");

            // Iterate through each paragraph in the first section of the document
            for (int j = 0; j < document.Sections[0].Paragraphs.Count; j++)
            {
                // Get the current paragraph from the first section
                Paragraph paragraph = document.Sections[0].Paragraphs[j];

                // Iterate through each child object within the current paragraph
                for (int i = 0; i < paragraph.ChildObjects.Count; i++)
                {
                    // Check if the current child object is a Shape (which can contain SmartArt)
                    if (paragraph.ChildObjects[i] is Spire.Doc.Fields.Shapes.Shape)
                    {
                        // Cast the child object to a Shape object
                        Spire.Doc.Fields.Shapes.Shape shape = paragraph.ChildObjects[i] as Spire.Doc.Fields.Shapes.Shape;

                        // Check if this shape contains a SmartArt graphic
                        if (shape.HasSmartArt)
                        {
                            // Remove the SmartArt shape from the paragraph's items collection
                            paragraph.Items.RemoveAt(i);

                            // Decrement the loop counter since we removed an item and the next item has shifted down
                            i--;
                        }
                    }
                }
            }

            // Define the output file name for saving the modified document
            string result = "RemoveSmartArt.docx";

            // Save the modified document (with SmartArt removed) to a file in Docx2016 format
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
