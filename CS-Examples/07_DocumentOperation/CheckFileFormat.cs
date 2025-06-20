using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;

namespace CheckFileFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
               // Define the input file path
				string input = @"..\..\..\..\..\..\Data\Template.docx";

				// Create a new Document object
				Document doc = new Document();

				// Load the document from the specified input file
				doc.LoadFromFile(input);

				// Get the detected format type of the document
				FileFormat ff = doc.DetectedFormatType;

				// Initialize a string to hold the file format information
				string fileFormat = "The file format is ";

				// Use a switch statement to determine the file format and update the fileFormat string accordingly
				switch (ff)
				{
					case FileFormat.Doc:
						fileFormat += "Microsoft Word 97-2003 document.";
						break;
					case FileFormat.Dot:
						fileFormat += "Microsoft Word 97-2003 template.";
						break;
					case FileFormat.Docx:
						fileFormat += "Office Open XML WordprocessingML Macro-Free Document.";
						break;
					case FileFormat.Docm:
						fileFormat += "Office Open XML WordprocessingML Macro-Enabled Document.";
						break;
					case FileFormat.Dotx:
						fileFormat += "Office Open XML WordprocessingML Macro-Free Template.";
						break;
					case FileFormat.Dotm:
						fileFormat += "Office Open XML WordprocessingML Macro-Enabled Template.";
						break;
					case FileFormat.Rtf:
						fileFormat += "RTF format.";
						break;
					case FileFormat.WordML:
						fileFormat += "Microsoft Word 2003 WordprocessingML format.";
						break;
					case FileFormat.Html:
						fileFormat += "HTML format.";
						break;
					case FileFormat.WordXml:
						fileFormat += "Microsoft Word XML format for Word 2007-2013.";
						break;
					case FileFormat.Odt:
						fileFormat += "OpenDocument Text.";
						break;
					case FileFormat.Ott:
						fileFormat += "OpenDocument Text Template.";
						break;
					case FileFormat.DocPre97:
						fileFormat += "Microsoft Word 6 or Word 95 format.";
						break;
					default:
						fileFormat += "Unknown format.";
						break;
				}

				// Display a message box with the file format information
				MessageBox.Show(fileFormat);

				// Dispose of the Document object to release resources
				doc.Dispose();
        }
    }
}
