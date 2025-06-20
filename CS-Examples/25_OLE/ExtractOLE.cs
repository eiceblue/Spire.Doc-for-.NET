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
using System.IO;

namespace ExtractOLE
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

			// Create a new document object
			Document doc = new Document();

			// Load the document from a file
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\OLEs.docx");

			// Iterate through each section in the document
			foreach (Section sec in doc.Sections)
			{
				// Iterate through each child object in the section's body
				foreach (DocumentObject obj in sec.Body.ChildObjects)
				{
					// Check if the object is a paragraph
					if (obj is Paragraph)
					{
						// Cast the object to a paragraph
						Paragraph par = obj as Paragraph;
						// Iterate through each child object in the paragraph
						foreach (DocumentObject o in par.ChildObjects)
						{
							// Check if the child object is an OLE object
							if (o.DocumentObjectType == DocumentObjectType.OleObject)
							{
								// Cast the object to a DocOleObject
								DocOleObject Ole = o as DocOleObject;
								// Get the type of the OLE object
								string s = Ole.ObjectType;

								// Perform actions based on the OLE object type
								if (s == "AcroExch.Document.DC")
								{
									// Save the OLE object as a PDF file
									File.WriteAllBytes("Result.pdf", Ole.NativeData);
									// Open the PDF file with the default file viewer
									FileViewer("Result.pdf");
								}
								else if (s == "Excel.Sheet.8")
								{
									// Save the OLE object as an Excel file
									File.WriteAllBytes("ExcelResult.xls", Ole.NativeData);
									// Open the Excel file with the default file viewer
									FileViewer("ExcelResult.xls");
								}
								else if (s == "PowerPoint.Show.12")
								{
									// Save the OLE object as a PowerPoint file
									File.WriteAllBytes("PPTResult.pptx", Ole.NativeData);
									// Open the PowerPoint file with the default file viewer
									FileViewer("PPTResult.pptx");
								}
							}
						}
					}
				}
			}

			// Dispose the document object
			doc.Dispose();
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
