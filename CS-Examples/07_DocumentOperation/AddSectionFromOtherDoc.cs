using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Text;

namespace AddSectionFromOtherDoc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class and load a document from the specified file path.
			Document TarDoc = new Document(@"..\..\..\..\..\..\..\Data\SampleB_1.docx");

			// Create a new instance of the Document class and load another document from the specified file path.
			Document SouDoc = new Document(@"..\..\..\..\..\..\..\Data\Sample_two sections.docx");

			// Get the second section from SouDoc.
			Section Ssection = SouDoc.Sections[1];

			// Clone the second section and add it to TarDoc.
			TarDoc.Sections.Add(Ssection.Clone());

			// Specify the output file name.
			string result = "result.docx";

			// Save the modified TarDoc to a file with the specified output file name and format (Docx).
			TarDoc.SaveToFile(result, FileFormat.Docx);

			// Clean up resources used by the TarDoc and SouDoc.
			TarDoc.Dispose();
			SouDoc.Dispose();

            //Launch result file
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
