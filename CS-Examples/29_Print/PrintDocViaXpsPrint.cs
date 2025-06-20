using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;

namespace PrintDocViaXpsPrint
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			// Create a new MemoryStream for storing the document as XPS
			using (MemoryStream ms = new MemoryStream())
			{
				// Instantiate a new Document object
				using (Document document = new Document())
				{
					// Load the Word document from the specified template file
					document.LoadFromFile(@"..\..\..\..\..\..\Data\Template.docx");

					// Save the document to the MemoryStream as XPS format
					document.SaveToStream(ms, FileFormat.XPS);
				}

				// Reset the position of the MemoryStream to the beginning
				ms.Position = 0;

				// Specify the printer name to be used for printing
				String printerName = "HP LaserJet P1007";

				// Print the XPS document using the specified printer and job name
				XpsPrint.XpsPrintHelper.Print(ms, printerName, "My printing job", true);
			}
        }
    }
}
