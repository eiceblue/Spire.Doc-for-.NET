using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetImageBackground
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object and load a Word document from a specific file path
            Document document = new Document(@"..\..\..\..\..\..\Data\Template.docx");

            // Set the background type of the document to picture
            document.Background.Type = BackgroundType.Picture;

            // Set the background picture of the document by loading an image from a file path
            document.Background.Picture = Image.FromFile(@"..\..\..\..\..\..\Data\Background.png");
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
             document.Background.SetPicture(inputFile_Img);
            */


            // Save the modified document to a file with the specified format (Docx)
            document.SaveToFile("ImageBackground.docx", FileFormat.Docx);

            // Dispose the Document object to release resources
            document.Dispose();

            //launching the Word file.
            WordDocViewer("ImageBackground.docx");


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
