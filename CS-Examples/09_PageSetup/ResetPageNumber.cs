using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ResetPageNumber
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object and load the first document file.
			Document document1 = new Document();
			document1.LoadFromFile(@"..\..\..\..\..\..\Data\ResetPageNumber1.docx");

			// Create a new Document object and load the second document file.
			Document document2 = new Document();
			document2.LoadFromFile(@"..\..\..\..\..\..\Data\ResetPageNumber2.docx");

			// Create a new Document object and load the third document file.
			Document document3 = new Document();
			document3.LoadFromFile(@"..\..\..\..\..\..\Data\ResetPageNumber3.docx");

			// Copy sections from document2 to document1.
			foreach (Section sec in document2.Sections)
			{
				document1.Sections.Add(sec.Clone());
			}

			// Copy sections from document3 to document1.
			foreach (Section sec in document3.Sections)
			{
				document1.Sections.Add(sec.Clone());
			}

			// Modify field types in footer sections of document1.
			foreach (Section sec in document1.Sections)
			{
				foreach (DocumentObject obj in sec.HeadersFooters.Footer.ChildObjects)
				{
					if (obj.DocumentObjectType == DocumentObjectType.StructureDocumentTag)
					{
						DocumentObject para = obj.ChildObjects[0];
						foreach (DocumentObject item in para.ChildObjects)
						{
							if (item.DocumentObjectType == DocumentObjectType.Field)
							{
								if ((item as Field).Type == FieldType.FieldNumPages)
								{
									// Change the field type to FieldSectionPages.
									(item as Field).Type = FieldType.FieldSectionPages;
								}
							}
						}
					}
				}
			}

			// Reset page numbering for specific sections in document1.
			document1.Sections[1].PageSetup.RestartPageNumbering = true;
			document1.Sections[1].PageSetup.PageStartingNumber = 1;
			document1.Sections[2].PageSetup.RestartPageNumbering = true;
			document1.Sections[2].PageSetup.PageStartingNumber = 1;

			// Specify the filename for the resulting document with reset page numbering.
			string result = "Result-ResetPageNumber.docx";

			// Save the modified document to a file in the Docx2013 format.
			document1.SaveToFile(result, FileFormat.Docx2013);

			// Release the resources associated with the documents.
			document1.Dispose();
			document2.Dispose();
			document3.Dispose();

            //Launch the MS Word file.
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
