using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ConvertToHtml
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class.
			Document document = new Document();

			// Load a Word document from the specified file path using relative path notation.
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\ToHtmlTemplate.docx");

			// Save the loaded document as an HTML file named "Sample.html".
			document.SaveToFile("Sample.html", FileFormat.Html);

			// Release system resources associated with the Document object.
			document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("Sample.html");
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
