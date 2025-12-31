using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ReplaceWithImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Define a string that contains the path to the input document
            string input = @"..\..\..\..\..\..\Data\Template.docx";

            // Create a new instance of the Document class
            Document doc = new Document();

            // Load the content of the document from the defined input path
            doc.LoadFromFile(input);
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            //Find the string "E-iceblue" in the document
            TextSelection[] selections = doc.FindAllString("E-iceblue", true, true);
            int index = 0;
            TextRange range = null;

            //Remove the text and replace it with Image
            foreach (TextSelection selection in selections)
            {
                DocPicture pic = new DocPicture(doc);
                pic.LoadImage(inputFile_2);

                range = selection.GetAsOneRange();
                index = range.OwnerParagraph.ChildObjects.IndexOf(range);
                range.OwnerParagraph.ChildObjects.Insert(index, pic);
                range.OwnerParagraph.ChildObjects.Remove(range);
            } 
            */




            // Define an Image object from a file located at "..\..\..\..\..\..\Data\E-iceblue.png"
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png");

            // Find all occurrences of the string "E-iceblue" in the document
            TextSelection[] selections = doc.FindAllString("E-iceblue", true, true);

            // Initialize an index to keep track of the current range's position
            int index = 0;

            // Initialize a TextRange to keep the current range of text being processed
            TextRange range = null;

            // Iterate over all the occurrences of "E-iceblue" found in the document
            foreach (TextSelection selection in selections)
            {
                // Create a new DocPicture object and load the defined image into it
                DocPicture pic = new DocPicture(doc);
                pic.LoadImage(image);

                // Get the current range of text being processed
                range = selection.GetAsOneRange();

                // Get the current index of the TextRange within its owner paragraph's ChildObjects collection
                index = range.OwnerParagraph.ChildObjects.IndexOf(range);

                // Insert the image into the owner paragraph's ChildObjects collection at the position of the TextRange
                range.OwnerParagraph.ChildObjects.Insert(index, pic);

                // Remove the TextRange from its owner paragraph's ChildObjects collection
                range.OwnerParagraph.ChildObjects.Remove(range);
            }

            // Define the output path and filename for the modified document
            string output = "ReplaceWithImage.docx";

            // Save the modified document to the defined output path with the .docx file format
            doc.SaveToFile(output, FileFormat.Docx);

            // Dispose of the Document object to release its resources
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
