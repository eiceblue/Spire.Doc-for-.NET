using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;
using Spire.Doc.Fields;

namespace GetContentControlProperty
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a new document and load from file
            string input = @"..\..\..\..\..\..\Data\ContentControl.docx"; ;
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get all structureTags in the Word document
            StructureTags structureTags = GetAllTags(doc);
            //Get all StructureDocumentTagInline objects
            List<StructureDocumentTagInline> tagInlines = structureTags.tagInlines;
            string property = null;
            property += "Alias of contentControl" + "\t" + "ID          " + "\t" + "Tag             " + "\t" + "STDType        " + "\r\n";
            //Get properties of all tagInlines
            for (int i = 0; i < tagInlines.Count; i++)
            {
                string alias = tagInlines[i].SDTProperties.Alias;
                decimal id = tagInlines[i].SDTProperties.Id;
                string tag = tagInlines[i].SDTProperties.Tag;
                string STDType = tagInlines[i].SDTProperties.SDTType.ToString();
                property += alias + ",\t" + id + ",\t" + tag + ",\t" + STDType + "\r\n";
            }

            //Get all StructureDocumentTag objects
            List<StructureDocumentTag> tags = structureTags.tags;
            //Get properties of all tags
            for (int i = 0; i < tags.Count; i++)
            {
                string alias = tags[i].SDTProperties.Alias;
                decimal id = tags[i].SDTProperties.Id;
                string tag = tags[i].SDTProperties.Tag;
                string STDType = tags[i].SDTProperties.SDTType.ToString();
                property += alias + ",\t" + id + ",\t" + tag + ",\t" + STDType + "\r\n";
            }

            //Save the property to a text document and launch it
            string output = "Property.txt";
            File.WriteAllText(output, property.ToString());
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        //Get all StructureTags of the Word document
        private static StructureTags GetAllTags(Document document)
        {
            StructureTags structureTags = new StructureTags();
            foreach (Section section in document.Sections)
            {
                foreach (DocumentObject obj in section.Body.ChildObjects)
                {
                    if (obj.DocumentObjectType == DocumentObjectType.StructureDocumentTag)
                    {
                        structureTags.tags.Add(obj as StructureDocumentTag);

                    }
                    else if (obj.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        foreach (DocumentObject pobj in (obj as Paragraph).ChildObjects)
                        {
                            if (pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline)
                            {
                                structureTags.tagInlines.Add(pobj as StructureDocumentTagInline);
                            }
                        }
                    }
                    else if (obj.DocumentObjectType == DocumentObjectType.Table)
                    {
                        foreach (TableRow row in (obj as Table).Rows)
                        {
                            foreach (TableCell cell in row.Cells)
                            {
                                foreach (DocumentObject cellChild in cell.ChildObjects)
                                {
                                    if (cellChild.DocumentObjectType == DocumentObjectType.StructureDocumentTag)
                                    {
                                        structureTags.tags.Add(cellChild as StructureDocumentTag);
                                    }
                                    else if (cellChild.DocumentObjectType == DocumentObjectType.Paragraph)
                                    {
                                        foreach (DocumentObject pobj in (cellChild as Paragraph).ChildObjects)
                                        {
                                            if (pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline)
                                            {
                                                structureTags.tagInlines.Add(pobj as StructureDocumentTagInline);
                                            }
                                        }
                                    }
                                }
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
            List<StructureDocumentTag> m_tags;
            public List<StructureDocumentTag> tags
            {
                get
                {
                    if (m_tags == null)
                        m_tags = new List<StructureDocumentTag>();
                    return m_tags;
                }
                set
                {
                    m_tags = value;
                }
            }
        }
    }
}
