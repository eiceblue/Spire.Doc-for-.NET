using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Interface;

namespace RemoveControlsKeepContent
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

            // Load a document file from a specified path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveControlsKeepContent.docx");

            // Iterate through the sections in the document
            for (int s = 0; s < doc.Sections.Count; s++)
            {
                // Get the current section
                Section section = doc.Sections[s];

                // Iterate through the child objects in the section's body
                for (int i = 0; i < section.Body.ChildObjects.Count; i++)
                {
                    // Check if the child object is a paragraph
                    if (section.Body.ChildObjects[i] is Paragraph)
                    {
                        // Get the paragraph object
                        Paragraph para = section.Body.ChildObjects[i] as Paragraph;

                        // Iterate through the child objects in the paragraph
                        for (int j = 0; j < para.ChildObjects.Count; j++)
                        {
                            // Check if the child object is a StructureDocumentTagInline
                            if (para.ChildObjects[j] is StructureDocumentTagInline)
                            {
                                // Get the StructureDocumentTagInline object
                                StructureDocumentTagInline sdt = para.ChildObjects[j] as StructureDocumentTagInline;

                                // Remove this control while retaining its content
                                sdt.RemoveSelfOnly();

                            }
                        }
                    }

                    // Check if the child object is a StructureDocumentTag
                    if (section.Body.ChildObjects[i] is StructureDocumentTag)
                    {
                        // Get the StructureDocumentTag object
                        StructureDocumentTag sdt = section.Body.ChildObjects[i] as StructureDocumentTag;

                        // Remove this control while retaining its content
                        sdt.RemoveSelfOnly();
                    }
                }
            }

            // Save the modified document to a new file
            string output = "RemoveControlsKeepContent_result.docx";
            doc.SaveToFile(output, FileFormat.Docx2016);

            // Dispose the document object
            doc.Dispose();

            WordDocViewer(output);
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
