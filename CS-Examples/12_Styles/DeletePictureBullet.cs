using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Documents;
using Spire.Doc.Interface;

namespace DeletePictureBullet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // Initialize a new Document object to work with Word documents.
            Document doc = new Document();

            // Load an existing Word document from the specified file path into the document object.
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\PictureBullets.docx");

            // Access the collection of list definitions (styles) within the document.
            ListCollection lists = doc.ListReferences;

            // Get the second level (index 1) of the first list definition (index 0) in the document.
            ListLevel listLevel = lists[0].Levels[1];

            // Remove any picture bullet associated with this list level.
            listLevel.DeletePictureBullet();

            // Save the modified document to a new file named "DeletePictureBullet.docx" in Word 2016 format.

            doc.SaveToFile("DeletePictureBullet.docx", FileFormat.Docx2016);
            // Close the document to release internal resources such as file handles.

            doc.Close();

            // Explicitly dispose of the document object to ensure memory is freed immediately.
            doc.Dispose();

            //Launching the Word file.
            WordDocViewer("DeletePictureBullet.docx");
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
