using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Fields;
namespace AddCoverImage
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
			Document doc = new Document();

			// Load a document from the specified file path.
			doc.LoadFromFile(@"..\..\..\..\..\..\..\Data\ToEpub.doc");

			// Create a new DocPicture object with the document as its owner.
			DocPicture picture = new DocPicture(doc);

			// Load an image from the specified file path and assign it to the picture.
			picture.LoadImage(Image.FromFile(@"..\..\..\..\..\..\..\Data\Cover.png"));

			// Specify the output file name for the EPUB file.
			string result = "AddCoverImage.epub";

			// Save the document as an EPUB file, including the cover image.
			doc.SaveToEpub(result, picture);

			// Dispose the document object to release resources.
			doc.Dispose();
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
