using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetGradientBackground
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object
			Document document = new Document();

			// Load a Word document from a specific file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_2.docx");

			// Set the background type of the document to gradient
			document.Background.Type = BackgroundType.Gradient;

			// Get the BackgroundGradient object of the document's background
			BackgroundGradient Test = document.Background.Gradient;

			// Set the first color of the gradient background to white
			Test.Color1 = Color.White;

			// Set the second color of the gradient background to light blue
			Test.Color2 = Color.LightBlue;

			// Set the shading variant of the gradient background to ShadingDown
			Test.ShadingVariant = GradientShadingVariant.ShadingDown;

			// Set the shading style of the gradient background to Horizontal
			Test.ShadingStyle = GradientShadingStyle.Horizontal;

			// Specify the output file name
			string result = "Result-SetGradientBackground.docx";

			// Save the modified document to a file with the specified format (Docx2013)
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose the Document object to release resources
			document.Dispose();

            //Launch the MS Word file.
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
