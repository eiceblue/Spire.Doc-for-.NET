using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertOLEAsIconViaStream
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			// Specify the output file name
			string output = "InsertOLEAsIconViaStream.docx";

			// Create a new document object
			Document doc = new Document();

			// Add a section to the document
			Section sec = doc.AddSection();

			// Add a paragraph to the section
			Paragraph par = sec.AddParagraph();

			// Open a stream for the OLE object data from the specified file
			Stream stream = File.OpenRead(@"..\..\..\..\..\..\Data\example.zip");

			// Create a DocPicture object and load an image from file
			DocPicture picture = new DocPicture(doc);
			Image image = Image.FromFile(@"..\..\..\..\..\..\Data\example.png");
			picture.LoadImage(image);

			// Append an OLE object to the paragraph using the provided stream, picture, and object type ("zip")
			DocOleObject obj = par.AppendOleObject(stream, picture, "zip");

			// Set the OLE object to be displayed as an icon
			obj.DisplayAsIcon = true;

			// Save the document to a file in Docx2013 format
			doc.SaveToFile(output, FileFormat.Docx2013);

			// Dispose the document object
			doc.Dispose();

            //Launching the Word file.
            WordDocViewer(output);


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
