using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

namespace TextWaterMark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
       
            // Load the document from a template file
            Document document = new Document(@"..\..\..\..\..\..\Data\Template.docx");

            // Insert text watermark into the first section of the document
            InsertTextWatermark(document.Sections[0]);

            // Save the modified document to a new file
            document.SaveToFile("Sample.docx", FileFormat.Docx);

            // Dispose the document object
            document.Dispose();

            //Launching the Word file.
            WordDocViewer("Sample.docx");


        }
		private void InsertTextWatermark(Section section) {
			// Create a TextWatermark object
			TextWatermark txtWatermark = new TextWatermark();
			// Set the text for the watermark
			txtWatermark.Text = "E-iceblue";
			// Set the font size of the watermark
			txtWatermark.FontSize = 95;
			// Set the color of the watermark
			txtWatermark.Color = Color.Blue;
			// Set the layout of the watermark
			txtWatermark.Layout = WatermarkLayout.Diagonal;
			// Set the watermark for the document section
			section.Document.Watermark = txtWatermark;
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
