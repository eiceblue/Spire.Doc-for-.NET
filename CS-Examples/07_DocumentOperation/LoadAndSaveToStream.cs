using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;
using Spire.Doc.Fields;
namespace LoadAndSaveToStream
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Define the input file path using a relative path
			string input = @"..\..\..\..\..\..\Data\Template.docx";

			// Open the input file in read mode and obtain a Stream object
			Stream stream = File.OpenRead(input);

			// Create a new instance of the Document class by loading the document from the input stream
			Document doc = new Document(stream);

			// Close the input stream to release resources
			stream.Close();

			// Perform operations on the document

			// Create a new MemoryStream to store the document
			MemoryStream newStream = new MemoryStream();

			// Save the document to the new memory stream in RTF format
			doc.SaveToStream(newStream, FileFormat.Rtf);

			// Reset the position of the memory stream to the beginning
			newStream.Position = 0;

			// Specify the output file name
			string result = "LoadAndSaveToStream_out.rtf";

			// Write the contents of the memory stream to a file with the specified output file name
			File.WriteAllBytes(result, newStream.ToArray());

			// Dispose of the document object to free up resources
			doc.Dispose();

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
