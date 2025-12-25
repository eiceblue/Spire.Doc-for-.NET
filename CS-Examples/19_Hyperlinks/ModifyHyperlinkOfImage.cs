using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ModifyHyperlinkOfImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Input file path
            String input = "..\\..\\..\\..\\..\\..\\Data\\ImageHyperlink.docx";

            //Output file path
            String output = "ModifyHyperlinkOfImage_output.docx";

            //Create word document
            Document document = new Document();

            //Load a document
            document.LoadFromFile(input);

            // Iterate through each section in the document
            foreach (Section section in document.Sections)
            {
                // Iterate through each paragraph within the current section
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    // Iterate through each child object within the current paragraph
                    foreach (DocumentObject documentObject in paragraph.ChildObjects)
                    {
                        // Check if the current document object is a picture (DocPicture)
                        if (documentObject is DocPicture)
                        {
                            // Cast the document object to a DocPicture type
                            DocPicture pic = documentObject as DocPicture;

                            // Check if the picture has a hyperlink associated with it
                            if (pic.HasHyperlink)
                            {
                                // Update the hyperlink of the picture to a new URL
                                pic.HRef = "https://www.e-iceblue.com/Introduce/word-for-net-introduce.html";
                            }
                        }
                    }
                }
            }


            // Save to file
            document.SaveToFile(output, FileFormat.Docx2019);

            //Dispose the document
            document.Dispose();

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
