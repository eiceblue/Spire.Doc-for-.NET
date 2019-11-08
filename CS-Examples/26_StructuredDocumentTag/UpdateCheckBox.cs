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
            //Create a document
            Document document = new Document();

            //Load the document from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\CheckBoxContentControl.docx");

            //Call StructureTags
            StructureTags structureTags = GetAllTags(document);

            //Create list 
            List<StructureDocumentTagInline> tagInlines = structureTags.tagInlines;

            //Get the controls
            for (int i = 0; i < tagInlines.Count; i++)
            {
                //Get the type
                string type = tagInlines[i].SDTProperties.SDTType.ToString();

                //Update the status
                if (type == "CheckBox")
                {
                    SdtCheckBox scb = tagInlines[i].SDTProperties.ControlProperties as SdtCheckBox;
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
            //Save the document.
            document.SaveToFile("Output.docx", FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer("Output.docx");

        }

        static StructureTags GetAllTags(Document document)
        {

            //Create StructureTags
            StructureTags structureTags = new StructureTags();

            //Travel document sections
            foreach (Section section in document.Sections)
            {
                foreach (DocumentObject obj in section.Body.ChildObjects)
                {
                    //Travel document paragraphs
                    if (obj.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        foreach (DocumentObject pobj in (obj as Paragraph).ChildObjects)
                        {
                            //Get StructureDocumentTagInline
                            if (pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline)
                            {
                                structureTags.tagInlines.Add(pobj as StructureDocumentTagInline);
                            }
                        }
                    }

                }
            }
            return structureTags;
        }
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
