using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace InsertOleObjects
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            String dataPath = @"..\..\..\..\..\..\Data\Sample.docx";
            String imagePath = @"..\..\..\..\..\..\Data\E-iceblue.png";
            String progId = "Word.Document";


            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            SKBitmap image = SKBitmap.Decode(imagePath);
            */


            // Load an image from the specified file path to be used as the OLE object icon.
            Image image = Image.FromFile(imagePath);

            // Create a new document instance.
            Document doc = new Document();

            // Create a document navigator to facilitate content insertion into the document.
            DocumentNavigator navigator = new DocumentNavigator(doc);

            // Insert a line of text explaining that the following OLE object is a linked object displayed as an icon.
            navigator.Writeln("The object is embedded, displayed as an icon.");

            // Insert an OLE object as an embedded object, displayed as an icon, using the specified data file, program ID, and image.
            navigator.InsertOleObject(dataPath, progId, false, true, image);

            // Insert a blank line for spacing.
            navigator.Writeln();

            // Insert explanatory text indicating the next OLE object is a linked object but not shown as an icon.
            navigator.Writeln("The object is a link displayed as an icon, specifying a ProgID.");

            // Insert an OLE object as a linked object (not embedded), not displayed as an icon.
            navigator.InsertOleObject(dataPath, progId, true, true, image);

            // Insert another blank line for spacing.
            navigator.Writeln();

            // Insert text describing another linked OLE object shown as an icon.
            navigator.Writeln("The object is a link displayed as an icon, without specifying a ProgID.");

            // Insert an OLE object as a linked object (without specifying a ProgID), displayed as an icon.
            navigator.InsertOleObject(dataPath, true, true, image);

            // Insert another blank line for spacing.
            navigator.Writeln();

            // Insert text indicating the next OLE object uses a stream as its data source and is displayed as an icon.
            navigator.Writeln("The stream of the Ole data source is displayed as an icon.");

            // Open the OLE data file as a read-only file stream.
            FileStream oleStream = new FileStream(dataPath, FileMode.Open, FileAccess.Read);

            // Insert an OLE object using the file stream, with the given ProgID, embedded (not linked), and displayed as an icon.
            navigator.InsertOleObject(oleStream, progId, true, image);

            // Save the final document to a DOCX file.
            doc.SaveToFile("InsertOleObjects.docx", FileFormat.Docx);

            // Close the document to release internal resources.
            doc.Close();

            // Explicitly dispose of the document object to free memory immediately.
            doc.Dispose();

            //Launch the MS Word file.
            WordDocViewer("InsertOleObjects.docx");
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
