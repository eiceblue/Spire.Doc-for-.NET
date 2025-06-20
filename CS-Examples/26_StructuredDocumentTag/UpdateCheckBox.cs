using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;

namespace UpdateCheckBox
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
            Document document = new Document();

            // Load a document from the specified file path
            document.LoadFromFile(@"..\..\..\..\..\..\Data\CheckBoxContentControl.docx");

            // Get all the StructureTags from the document
            StructureTags structureTags = GetAllTags(document);

            // Get the list of StructureDocumentTagInline objects from the StructureTags
            List<StructureDocumentTagInline> tagInlines = structureTags.tagInlines;

            // Iterate through the list of StructureDocumentTagInline objects
            for (int i = 0; i < tagInlines.Count; i++)
            {
                // Get the SDTType of the current StructureDocumentTagInline
                string type = tagInlines[i].SDTProperties.SDTType.ToString();

                // Check if the SDTType is "CheckBox"
                if (type == "CheckBox")
                {
                    // Get the SdtCheckBox from the ControlProperties of the StructureDocumentTagInline
                    SdtCheckBox scb = tagInlines[i].SDTProperties.ControlProperties as SdtCheckBox;

                    // Toggle the Checked property of the SdtCheckBox
                    if (scb.Checked)
                    {
                        scb.Checked = false;
                    }
                    else
                    {
                        scb.Checked = true;
                    }
                }
            }

            // Save the modified document to "Output.docx" in DOCX format
            document.SaveToFile("Output.docx", FileFormat.Docx);

            // Dispose the document object
            document.Dispose();

            //Launch the Word file.
            WordDocViewer("Output.docx");

        }

        // Define a method named "GetAllTags" that takes a Document object as input and returns a StructureTags object
        static StructureTags GetAllTags(Document document)
        {
            // Create a new StructureTags object to store the StructureDocumentTagInline objects
            StructureTags structureTags = new StructureTags();

            // Iterate through the sections in the document
            foreach (Section section in document.Sections)
            {
                // Iterate through the child objects in the section's body
                foreach (DocumentObject obj in section.Body.ChildObjects)
                {
                    // Check if the current object is a Paragraph
                    if (obj.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        // Iterate through the child objects in the paragraph
                        foreach (DocumentObject pobj in (obj as Paragraph).ChildObjects)
                        {
                            // Check if the current object is a StructureDocumentTagInline
                            if (pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline)
                            {
                                // Add the StructureDocumentTagInline to the tagInlines list in the StructureTags object
                                structureTags.tagInlines.Add(pobj as StructureDocumentTagInline);
                            }
                        }
                    }
                }
            }

            // Return the StructureTags object containing the collected StructureDocumentTagInline objects
            return structureTags;
        }

        // Define a public class named "StructureTags"
        public class StructureTags
        {
            List<StructureDocumentTagInline> m_tagInlines;
            public List<StructureDocumentTagInline> tagInlines
            {
                get
                {
                    if (m_tagInlines == null)
                        m_tagInlines = new List<StructureDocumentTagInline>();
                    return m_tagInlines;
                }
                set
                {
                    m_tagInlines = value;
                }
            }
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
