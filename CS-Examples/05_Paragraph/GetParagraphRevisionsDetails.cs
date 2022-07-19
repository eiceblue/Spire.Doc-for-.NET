using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace GetParagraphRevisionsDetails
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load word document
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Revisions.docx");

            StringBuilder builder = new StringBuilder();

            //loop paragraph
            foreach (Section section in document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    if (paragraph.IsDeleteRevision)
                    {
                        builder.AppendLine(string.Format("The section {0} paragraph {1} has been changed (deleted).", document.GetIndex(section), section.GetIndex(paragraph)));
                        builder.AppendLine("Author: " + paragraph.DeleteRevision.Author);
                        builder.AppendLine("DateTime: " + paragraph.DeleteRevision.DateTime);
                        builder.AppendLine("Type: " + paragraph.DeleteRevision.Type);
                        builder.AppendLine("");
                    }
                    else if (paragraph.IsInsertRevision)
                    {
                        builder.AppendLine(string.Format("The section {0} paragraph {1} has been changed (inserted).", document.GetIndex(section), section.GetIndex(paragraph)));
                        builder.AppendLine("Author: " + paragraph.InsertRevision.Author);
                        builder.AppendLine("DateTime: " + paragraph.InsertRevision.DateTime);
                        builder.AppendLine("Type: " + paragraph.InsertRevision.Type);
                        builder.AppendLine("");
                    }
                    else
                    {
                        foreach (DocumentObject obj in paragraph.ChildObjects)
                        {
                            if (obj.DocumentObjectType.Equals(DocumentObjectType.TextRange))
                            {
                                TextRange textRange = obj as TextRange;
                                {
                                    if (textRange.IsDeleteRevision)
                                    {
                                        builder.AppendLine(string.Format("The section {0} paragraph {1} textrange {2} has been changed (deleted).", document.GetIndex(section), section.GetIndex(paragraph), paragraph.GetIndex(textRange)));
                                        builder.AppendLine("Author: " + textRange.DeleteRevision.Author);
                                        builder.AppendLine("DateTime: " + textRange.DeleteRevision.DateTime);
                                        builder.AppendLine("Type: " + textRange.DeleteRevision.Type);
                                        builder.AppendLine("Change Text: " + textRange.Text);
                                        builder.AppendLine("");
                                    }
                                    else if (textRange.IsInsertRevision)
                                    {
                                        builder.AppendLine(string.Format("The section {0} paragraph {1} textrange {2} has been changed (inserted).", document.GetIndex(section), section.GetIndex(paragraph), paragraph.GetIndex(textRange)));
                                        builder.AppendLine("Author: " + textRange.InsertRevision.Author);
                                        builder.AppendLine("DateTime: " + textRange.InsertRevision.DateTime);
                                        builder.AppendLine("Type: " + textRange.InsertRevision.Type);
                                        builder.AppendLine("Change Text: " + textRange.Text);
                                        builder.AppendLine("");
                                    }
                                }
                            }
                        }
                    }
                }

                //Write the contents in a TXT file
                string output = "GetParagraphRevisionsDetails.txt";
                File.WriteAllText(output, builder.ToString());

                //Launch the file
                TxtViewer(output);
            }
        }

        private void TxtViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }


    }
}
