using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using System.IO;

namespace GetText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {   
			// Create a new instance of the Document class and load a Word document from the specified file path.
			Document document = new Document(@"..\..\..\..\..\..\Data\ExtractText.docx");

			// Extract the text content from the document.
			string text = document.GetText();

			// Write the extracted text to a text file named "Extract.txt".
			File.WriteAllText("Extract.txt", text);

			// Clean up resources used by the document.
			document.Dispose();

            //launch the file.
            WordDocViewer("Extract.txt");
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
