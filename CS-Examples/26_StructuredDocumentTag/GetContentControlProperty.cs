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
 
			// Specify the input file path
			string input = @"..\..\..\..\..\..\Data\ContentControl.docx";

			// Create a new document object
			Document doc = new Document();

			// Load the document from the specified file path
			doc.LoadFromFile(input);

			// Get all the structure tags in the document
			StructureTags structureTags = GetAllTags(doc);

			// Initialize variables for storing tag properties
			string alias = null;
			decimal id = 0;
			string tag = null;
			string property = "Alias of contentControl" + "\t" + "ID          " + "\t" + "Tag             " + "\t" + "STDType        " + "\r" + "Content        " + "\r\n";
			string sdtType = null;
			Paragraph paragraph = null;
			SdtType sdt = SdtType.RichText;
			string content = "";
			TextRange textRange = null;

			// Retrieve structure document tags and process their properties and content
			List<StructureDocumentTag> tags = structureTags.tags;
			for (int i = 0; i < tags.Count; i++)
			{
				alias = tags[i].SDTProperties.Alias;
				id = tags[i].SDTProperties.Id;
				tag = tags[i].SDTProperties.Tag;
				sdt = tags[i].SDTProperties.SDTType;
				sdtType = sdt.ToString();
				if (sdt == SdtType.RichText || sdt == SdtType.Text)
				{
					if (tags[i].ChildObjects.Count > 0)
					{
						foreach (DocumentObject obj in tags[i].ChildObjects)
						{
							if (obj is Paragraph)
							{
								paragraph = obj as Paragraph;
								content += paragraph.Text;
							}
						}
					}
				}
				property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
				content = "";
			}

			// Retrieve structure document tag inlines and process their properties and content
			List<StructureDocumentTagInline> tagInlines = structureTags.tagInlines;
			for (int i = 0; i < tagInlines.Count; i++)
			{
				alias = tagInlines[i].SDTProperties.Alias;
				id = tagInlines[i].SDTProperties.Id;
				tag = tagInlines[i].SDTProperties.Tag;
				sdt = tagInlines[i].SDTProperties.SDTType;
				sdtType = sdt.ToString();
				if (sdt == SdtType.RichText || sdt == SdtType.Text)
				{
					if (tagInlines[i].ChildObjects.Count > 0)
					{
						foreach (DocumentObject obj in tagInlines[i].ChildObjects)
						{
							if (obj is TextRange)
							{
								textRange = obj as TextRange;
								content += textRange.Text;
							}
						}
					}
				}
				property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
				content = "";
			}

			// Retrieve structure document tag rows and process their properties and content
			List<StructureDocumentTagRow> rowTags = structureTags.rowTags;
			for (int i = 0; i < rowTags.Count; i++)
			{
				alias = rowTags[i].SDTProperties.Alias;
				id = rowTags[i].SDTProperties.Id;
				tag = rowTags[i].SDTProperties.Tag;
				sdt = rowTags[i].SDTProperties.SDTType;
				sdtType = sdt.ToString();
				if (sdt == SdtType.RichText || sdt == SdtType.Text)
				{
					if (rowTags[i].ChildObjects.Count > 0)
					{
						foreach (DocumentObject obj in rowTags[i].ChildObjects)
						{
							if (obj is Paragraph)
							{
								paragraph = obj as Paragraph;
								content += paragraph.Text;
							}
						}
					}
				}
				property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
				content = "";
			}

			// Retrieve structure document tag cells and process their properties and content
			List<StructureDocumentTagCell> cellTags = structureTags.cellTags;
			for (int i = 0; i < cellTags.Count; i++)
			{
				alias = cellTags[i].SDTProperties.Alias;
				id = cellTags[i].SDTProperties.Id;
				tag = cellTags[i].SDTProperties.Tag;
				sdt = cellTags[i].SDTProperties.SDTType;
				sdtType = sdt.ToString();
				if (sdt == SdtType.RichText || sdt == SdtType.Text)
				{
					if (cellTags[i].ChildObjects.Count > 0)
					{
						foreach (DocumentObject obj in cellTags[i].ChildObjects)
						{
							if (obj is Paragraph)
							{
								paragraph = obj as Paragraph;
								content += paragraph.Text;
							}
						}
					}
				}
				property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
				content = "";
			}

			// Specify the output file name
			string output = "Property.txt";

			// Write the property string to the output file
			File.WriteAllText(output, property.ToString());

			// Dispose the document object
			doc.Dispose();
			
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
                            if (row is StructureDocumentTagRow)
                            {
                                structureTags.rowTags.Add(row as StructureDocumentTagRow);
                            }
                            foreach (TableCell cell in row.Cells)
                            {
                                if (cell is StructureDocumentTagCell)
                                {
                                    structureTags.cellTags.Add(cell as StructureDocumentTagCell);
                                }
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
            List<StructureDocumentTagCell> m_celltags;
            public List<StructureDocumentTagCell> cellTags
            {
                get
                {
                    if (m_celltags == null)
                        m_celltags = new List<StructureDocumentTagCell>();
                    return m_celltags;
                }
                set
                {
                    m_celltags = value;
                }
            }
            List<StructureDocumentTagRow> m_rowTags;
            public List<StructureDocumentTagRow> rowTags
            {
                get
                {
                    if (m_rowTags == null)
                        m_rowTags = new List<StructureDocumentTagRow>();
                    return m_rowTags;
                }
                set
                {
                    m_rowTags = value;
                }
            }
        }
        
    }
}
