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
            String input = @"..\..\..\..\..\..\Data\Template.docx";
            Document doc = new Document();
            // Load the document from disk.
            doc.LoadFromFile(input);

            // Create a new memory stream.
            MemoryStream outStream = new MemoryStream();
            // Save the document to stream.
            doc.SaveToStream(outStream, FileFormat.Docx);

            // Convert the document to bytes.
            byte[] docBytes = outStream.ToArray();

            // The bytes are now ready to be stored/transmitted.

            // Now reverse the steps to load the bytes back into a document object.
            MemoryStream inStream = new MemoryStream(docBytes);

            // Load the stream into a new document object.
            Document newDoc = new Document(inStream);
        }
    }
}
