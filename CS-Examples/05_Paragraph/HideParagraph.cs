using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace HideParagraph
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object.
            Document document = new Document();

            // Load an existing Word document from the specified file path.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

            // Get the first Section of the document.
            Section sec = document.Sections[0];

            // Get the first Paragraph of the Section.
            Paragraph para = sec.Paragraphs[0];

            // Iterate over the child DocumentObjects in the Paragraph.
            foreach (DocumentObject obj in para.ChildObjects)
            {
                // Check if the child DocumentObject is a TextRange.
                if (obj is TextRange)
                {
                    // Convert the child DocumentObject to a TextRange.
                    TextRange range = obj as TextRange;

                    // Set the Hidden property of the TextRange's CharacterFormat to true, hiding the text.
                    range.CharacterFormat.Hidden = true;
                }
            }

            // Specify the file name for the resulting Word document.
            string result = "Result-HideWordParagraph.docx";

            // Save the Document object to a file in Docx format with compatibility mode set to Docx2013.
            document.SaveToFile(result, FileFormat.Docx2013);

            // Dispose the Document object.
            document.Dispose();

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
