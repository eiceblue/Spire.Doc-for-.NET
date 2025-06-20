using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Formatting.Revisions;
using Spire.Doc.Fields;
using System.IO;

namespace ModifyRevisionTime
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object
			Document document = new Document();

			// Load a Word document from a file
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\ModifyRevisionTime.docx");

			// Initialize index variables
			int index_insertRevision = 0;
			int index_deleteRevision = 0;

			// Specify the date string and format
			string dateString = "2023/3/1 00:00:00";
			string format = "yyyy/M/d HH:mm:ss";

			// Parse the date string into a DateTime object using the specified format
			DateTime date = DateTime.ParseExact(dateString, format, null);

			// Iterate through the sections in the document
			foreach (Section sec in document.Sections)
			{
				// Iterate through the child objects in the section's body
				foreach (DocumentObject docItem in sec.Body.ChildObjects)
				{
					// Check if the child object is a Paragraph
					if (docItem is Paragraph)
					{
						// Cast the child object to a Paragraph
						Paragraph para = (Paragraph)docItem;
						
						// Check if the paragraph contains an insert revision
						if (para.IsInsertRevision)
						{
							// Increment the insert revision index
							index_insertRevision++;
							
							// Get the InsertRevision object for the paragraph
							EditRevision insRevison = para.InsertRevision;
							
							// Set the DateTime property of the insert revision to the specified date
							insRevison.DateTime = date;
						}
						// Check if the paragraph contains a delete revision
						else if (para.IsDeleteRevision)
						{
							// Increment the delete revision index
							index_deleteRevision++;
							
							// Get the DeleteRevision object for the paragraph
							EditRevision delRevison = para.DeleteRevision;
							
							// Set the DateTime property of the delete revision to the specified date
							delRevison.DateTime = date;
						}
						
						// Iterate through the child objects in the paragraph
						foreach (DocumentObject obj in para.ChildObjects)
						{
							// Check if the child object is a TextRange
							if (obj is TextRange)
							{
								// Cast the child object to a TextRange
								TextRange textRange = (TextRange)obj;
								
								// Check if the text range contains an insert revision
								if (textRange.IsInsertRevision)
								{
									// Increment the insert revision index
									index_insertRevision++;
									
									// Get the InsertRevision object for the text range
									EditRevision insRevison = textRange.InsertRevision;
									
									// Set the DateTime property of the insert revision to the specified date
									insRevison.DateTime = date;
								}
								// Check if the text range contains a delete revision
								else if (textRange.IsDeleteRevision)
								{
									// Increment the delete revision index
									index_deleteRevision++;
									
									// Get the DeleteRevision object for the text range
									EditRevision delRevison = textRange.DeleteRevision;
									
									// Set the DateTime property of the delete revision to the specified date
									delRevison.DateTime = date;
								}
							}
						}
					}
				}
			}

			// Save the modified document to a new file
			document.SaveToFile("ModifyRevisionTime_out.docx", FileFormat.Docx);

			// Dispose the Document object
			document.Dispose();

            WordDocViewer("ModifyRevisionTime_out.docx");

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
