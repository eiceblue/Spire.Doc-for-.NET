using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace LockContentControlContent
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
			// Specify the HTML table string
            String htmlString = "<table style=\"width: 100 % \">"
                + "<tr><th> Number </th><th> Name </th ><th>Age</th ></tr>"
                + "<tr><td> 1 </td><td> Smith </td><td> 50 </td></tr>"
                + "<tr> <td> 2 </td><td> Jackson </td><td> 94 </td> </tr>"
                + "</table>";

			// Create a new document
			Document doc = new Document();

			// Add a section to the document
			Section section = doc.AddSection();

			// Add a paragraph to the section
			Paragraph paragraph = section.AddParagraph();

			// Append HTML content to the paragraph
			paragraph.AppendHTML(htmlString);

			// Create a StructureDocumentTag
			StructureDocumentTag sdt = new StructureDocumentTag(doc);

			// Add a new section to the document
			Section section2 = doc.AddSection();

			// Add the StructureDocumentTag to the section's body
			section2.Body.ChildObjects.Add(sdt);

			// Set the type of the StructureDocumentTag to RichText
			sdt.SDTProperties.SDTType = SdtType.RichText;

			// Iterate through the child objects in the first section's body
			foreach (DocumentObject obj in section.Body.ChildObjects)
			{
				// Check if the object is a table
				if (obj.DocumentObjectType == DocumentObjectType.Table)
				{
					// Clone and add the table to the StructureDocumentTag's content
					sdt.SDTContent.ChildObjects.Add(obj.Clone());
				}
			}

			// Lock the content editing settings of the StructureDocumentTag
			sdt.SDTProperties.LockSettings = LockSettingsType.ContentLocked;

			// Remove the first section from the document
			doc.Sections.Remove(section);

			// Save the modified document to a file
			string result = "LockContentEditProperty_result.docx";
			doc.SaveToFile(result, Spire.Doc.FileFormat.Docx2013);

			// Dispose the document object
			doc.Dispose();
			
			
            //View the document
            FileViewer(result);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
