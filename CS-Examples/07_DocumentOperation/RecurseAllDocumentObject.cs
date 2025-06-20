using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RecurseAllDocumentObject
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class
			Document document = new Document();

			// Load the document from the specified file path using a relative path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

			// Create a StringBuilder object to store the output string
			StringBuilder builder = new StringBuilder();

			// Iterate through each section in the document
			foreach (Section section in document.Sections)
			{
				// Get the index of the current section
				int sectionIndex = document.GetIndex(section);
				
				// Append a formatted string indicating the section index and its child objects
				builder.AppendLine(string.Format("section index {0} has following ChildObjects", sectionIndex));

				// Iterate through each child object in the section's body
				foreach (DocumentObject obj in section.Body.ChildObjects)
				{
					// Get the index and type of the current child object
					builder.AppendLine(string.Format("Index: {0}, ChildObject Type: {1}", section.Body.GetIndex(obj), obj.DocumentObjectType));
					
					// Check if the child object is a paragraph
					if (obj.DocumentObjectType.Equals(DocumentObjectType.Paragraph))
					{
						// Convert the child object to a Paragraph
						Paragraph paragraph = obj as Paragraph;
						
						// Append a formatted string indicating the paragraph index and its child objects
						builder.AppendLine(string.Format("\tParagraph index {0} has following ChildObjects", section.Body.GetIndex(paragraph)));
						
						// Iterate through each child object in the paragraph
						foreach (DocumentObject obj2 in paragraph.ChildObjects)
						{
							// Append a formatted string indicating the index and type of the child object
							builder.AppendLine(string.Format("\tIndex: {0}, ChildObject Type: {1}", paragraph.GetIndex(obj2), obj2.DocumentObjectType));
						}
					}
				}
				
				// Append a blank line to separate sections
				builder.AppendLine(" ");
			}

			// Write the contents of the StringBuilder to a text file
			File.WriteAllText("RecurseAllDocumentObject.txt", builder.ToString());

			// Dispose of the document object to free up resources
			document.Dispose();

            //Launching the Word file.
            TextViewer("RecurseAllDocumentObject.txt");


        }

        private void TextViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}
