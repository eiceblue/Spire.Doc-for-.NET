using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace InsertImageWithNavigator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String imgpath = @"..\..\..\..\..\..\Data\E-iceblue.png";
            // Create a new Word document instance.
            Document doc = new Document();

            // Initialize a DocumentNavigator to help navigate and manipulate the document content.
            DocumentNavigator navigator = new DocumentNavigator(doc);

            // Write a line of text into the document indicating that an image will be inserted directly.
            navigator.Writeln("Insert the picture directly:");

            // Insert an image at the current cursor position using the specified image file path.
            navigator.InsertImage(imgpath);

            // Add a new section to the document.
            Section section = doc.AddSection();

            // Move the cursor to the second section (index 1, since sections are zero-based).
            navigator.MoveToSection(1);

            // Write a line of text explaining that the next image will have its dimensions set.
            navigator.Writeln("Set the width and height of the image:");

            // Insert an image with specified width and height (both set to 100 pixels).
            navigator.InsertImage(imgpath, 100, 100);

            // Insert a page break to start content on a new page.
            navigator.InsertBreak(BreakType.PageBreak);

            // Write a line of text describing more advanced image positioning and formatting.
            navigator.Writeln("Set the width, height, offset, and wrapping style of the image:");

            // Insert an image with detailed positioning and formatting
            navigator.InsertImage(imgpath, HorizontalOrigin.LeftMarginArea, 100, VerticalOrigin.Paragraph, 50, 100, 100, TextWrappingStyle.Through);

            // Save the document to a file.
            doc.SaveToFile("InsertImageWithNavigator.docx", FileFormat.Docx2019);

            // Close the document to release resources.
            doc.Close();

            //Launch the MS Word file.
            WordDocViewer("InsertImageWithNavigator.docx");
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
