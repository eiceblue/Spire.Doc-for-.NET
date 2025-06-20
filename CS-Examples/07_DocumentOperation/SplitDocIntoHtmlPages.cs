using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;
using Spire.Doc.Fields;
namespace SplitDocIntoHtmlPages
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
			String input = @"..\..\..\..\..\..\..\Data\SplitDocIntoHtmlPages.doc";

			// Define the output directory path
			string outDir = Path.Combine("output");

			// Create the output directory if it doesn't exist
			Directory.CreateDirectory(outDir);

			// Split the document into multiple HTML pages
			SplitDocIntoMultipleHtml(input, outDir);
        }
			// Split the document into multiple HTML pages
			private static void SplitDocIntoMultipleHtml(String input, string outDirectory)
			{
				// Load the document
				Document document = new Document();
				document.LoadFromFile(input);
				
				// Variable to hold the sub-document
				Document subDoc = null; 
				
				// Flag to check if it's the first element in the sub-document
				bool first = true; 
				
				// Index for naming the output HTML files
				int index = 0; 

				// Iterate through sections in the document
				foreach (Section sec in document.Sections)
				{
					// Iterate through elements in the section
					foreach (DocumentObject element in sec.Body.ChildObjects)
					{
						// Check if the element should be in the next document
						if (IsInNextDocument(element))
						{
							if (!first)
							{
								// Save the previous sub-document as an HTML file
								subDoc.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal;
								subDoc.HtmlExportOptions.ImageEmbedded = true;
								subDoc.SaveToFile(Path.Combine(outDirectory, String.Format("out-{0}.html", index++)), FileFormat.Html);
								subDoc = null;
							}
							first = false;
						}

						// Create a new sub-document if it doesn't exist
						if (subDoc == null)
						{
							subDoc = new Document();
							subDoc.AddSection();
						}

						// Add the element to the sub-document
						subDoc.Sections[0].Body.ChildObjects.Add(element.Clone());
					}
				}

				// Save the last sub-document as an HTML file, if it exists
				if (subDoc != null)
				{
					subDoc.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal;
					subDoc.HtmlExportOptions.ImageEmbedded = true;
					subDoc.SaveToFile(Path.Combine(outDirectory, String.Format("out-{0}.html", index++)), FileFormat.Html);
				}
			}
	
			// Check if the document element should be in the next document
			private static bool IsInNextDocument(DocumentObject element)
			{
				if (element is Paragraph)
				{
					Paragraph p = element as Paragraph;
					if (p.StyleName == "Heading1")
					{
						return true;
					}
				}
				return false;
			}
    }
}
