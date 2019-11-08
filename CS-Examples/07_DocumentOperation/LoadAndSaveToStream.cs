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
            String input = @"..\..\..\..\..\..\Data\Template.docx";

            // Open the stream. Read only access is enough to load a document.
            Stream stream = File.OpenRead(input);

            // Load the entire document into memory.
            Document doc = new Document(stream);

            // You can close the stream now, it is no longer needed because the document is in memory.
            stream.Close();
            // Do something with the document

            // Convert the document to a different format and save to stream.
            MemoryStream newStream = new MemoryStream();
            doc.SaveToStream(newStream, FileFormat.Rtf);

            // Rewind the stream position back to zero so it is ready for the next reader.
            newStream.Position = 0;

            // Save the document from stream, to disk. Normally you would do something with the stream directly,
            // For example, writing the data to a database.
            String result = "LoadAndSaveToStream_out.rtf";
            File.WriteAllBytes(result, newStream.ToArray());

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
