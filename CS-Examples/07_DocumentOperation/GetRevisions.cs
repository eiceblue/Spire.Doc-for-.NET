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

namespace GetRevisions
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
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\GetRevisions.docx");

			// Create a StringBuilder to store inserted revisions
			StringBuilder insertRevision = new StringBuilder();
			insertRevision.AppendLine("Insert revisions:");
			int index_insertRevision = 0;

			// Create a StringBuilder to store deleted revisions
			StringBuilder deleteRevision = new StringBuilder();
			deleteRevision.AppendLine("Delete revisions:");
			int index_deleteRevision = 0;

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
							insertRevision.AppendLine("Index: " + index_insertRevision);
							
							// Get the InsertRevision object for the paragraph
							EditRevision insRevison = para.InsertRevision;

							// Get the type of the insert revision
							EditRevisionType insType = insRevison.Type;
							insertRevision.AppendLine("Type: " + insType);
							
							// Get the author of the insert revision
							string insAuthor = insRevison.Author;
							insertRevision.AppendLine("Author: " + insAuthor);
						}
						// Check if the paragraph contains a delete revision
						else if (para.IsDeleteRevision)
						{
							// Increment the delete revision index
							index_deleteRevision++;
							deleteRevision.AppendLine("Index: " + index_deleteRevision);
							
							// Get the DeleteRevision object for the paragraph
							EditRevision delRevison = para.DeleteRevision;
							
							// Get the type of the delete revision
							EditRevisionType delType = delRevison.Type;
							deleteRevision.AppendLine("Type: " + delType);
							
							// Get the author of the delete revision
							string delAuthor = delRevison.Author;
							deleteRevision.AppendLine("Author: " + delAuthor);
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
									insertRevision.AppendLine("Index: " + index_insertRevision);
									
									// Get the InsertRevision object for the text range
									EditRevision insRevison = textRange.InsertRevision;
									
									// Get the type of the insert revision
									EditRevisionType insType = insRevison.Type;
									insertRevision.AppendLine("Type: " + insType);
									
									// Get the author of the insert revision
									string insAuthor = insRevison.Author;
									insertRevision.AppendLine("Author: " + insAuthor);
								}
								// Check if the text range contains a delete revision
								else if (textRange.IsDeleteRevision)
								{
									// Increment the delete revision index
									index_deleteRevision++;
									deleteRevision.AppendLine("Index: " + index_deleteRevision);
									
									// Get the DeleteRevision object for the text range
									EditRevision delRevison = textRange.DeleteRevision;
									
									// Get the type of the delete revision
									EditRevisionType delType = delRevison.Type;
									deleteRevision.AppendLine("Type: " + delType);
									
									// Get the author of the delete revision
									string delAuthor = delRevison.Author;
									deleteRevision.AppendLine("Author: " + delAuthor);
								}
							}
						}
					}
				}
			}

			// Write the inserted revisions to a text file
			File.WriteAllText("insertRevisions.txt", insertRevision.ToString());

			// Write the deleted revisions to a text file
			File.WriteAllText("deleteRevisions.txt", deleteRevision.ToString());

			// Dispose the Document object
			document.Dispose();
        }
    }
}
