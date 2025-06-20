using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;

namespace GetHeightAndWidthOfText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class.
			Document document = new Document();

			// Load a Word document from a specified file path.
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_2.docx");

			// Specify the text to search for within the document.
			string text = "Your Office Development Master";

			// Find the first occurrence of the specified text within the document.
			// Perform a case-insensitive search and include whole word matches.
			TextSelection selection = document.FindString(text, true, true);

			// Get the font used for the found text range.
			Font font = selection.GetAsOneRange().CharacterFormat.Font;

			// Create a fake image with dimensions of 1x1 pixels.
			Image fakeImage = new Bitmap(1, 1);
			Graphics graphics = Graphics.FromImage(fakeImage);

			// Measure the size (height and width) of the specified text using the font.
			SizeF size = graphics.MeasureString(text, font);

			// Create a StringBuilder to hold the content of the resulting text file.
			StringBuilder content = new StringBuilder();
			content.AppendLine("text height: " + size.Height);
			content.AppendLine("text width: " + size.Width);

			// Specify the file name for the resulting text file.
			string result = "Result-GetHeightAndWidthOfText.txt";

			// Write the content of the StringBuilder to the specified text file.
			File.WriteAllText(result, content.ToString());

			// Clean up resources used by the document.
			document.Dispose();

            //Launch the file.
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
