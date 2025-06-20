using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;

namespace ConvertDocToByte
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Define the input file path relative to the current directory.
			string input = @"..\..\..\..\..\..\Data\Template.docx";

			// Create a new instance of the Document class.
			Document doc = new Document();

			// Load the document from the specified input file.
			doc.LoadFromFile(input);

			// Create a new MemoryStream to store the document content.
			MemoryStream outStream = new MemoryStream();

			// Save the document to the MemoryStream in Docx format.
			doc.SaveToStream(outStream, FileFormat.Docx);

			// Convert the content of the MemoryStream to a byte array.
			byte[] docBytes = outStream.ToArray();

			// The bytes are now ready to be stored/transmitted.

			// Create a new MemoryStream from the byte array.
			MemoryStream inStream = new MemoryStream(docBytes);

			// Create a new Document object by loading from the MemoryStream.
			Document newDoc = new Document(inStream);

			// Dispose the existing document object.
			doc.Dispose();
        }
    }
}
