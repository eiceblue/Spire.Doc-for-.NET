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
            // Create a new Document object.
            Document document = new Document();

            // Load an existing Word document from the specified file path.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Revisions.docx");

            // Create a StringBuilder object to store the output details.
            StringBuilder builder = new StringBuilder();

            // Iterate over the Sections in the document.
            foreach (Section section in document.Sections)
            {
                // Iterate over the Paragraphs in each Section.
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    // Check if the Paragraph is a deleted revision.
                    if (paragraph.IsDeleteRevision)
                    {
                        // Append information about the deleted revision to the StringBuilder.
                        builder.AppendLine(string.Format("The section {0} paragraph {1} has been changed (deleted).", document.GetIndex(section), section.GetIndex(paragraph)));
                        builder.AppendLine("Author: " + paragraph.DeleteRevision.Author);
                        builder.AppendLine("DateTime: " + paragraph.DeleteRevision.DateTime);
                        builder.AppendLine("Type: " + paragraph.DeleteRevision.Type);
                        builder.AppendLine("");
                    }
                    // Check if the Paragraph is an inserted revision.
                    else if (paragraph.IsInsertRevision)
                    {
                        // Append information about the inserted revision to the StringBuilder.
                        builder.AppendLine(string.Format("The section {0} paragraph {1} has been changed (inserted).", document.GetIndex(section), section.GetIndex(paragraph)));
                        builder.AppendLine("Author: " + paragraph.InsertRevision.Author);
                        builder.AppendLine("DateTime: " + paragraph.InsertRevision.DateTime);
                        builder.AppendLine("Type: " + paragraph.InsertRevision.Type);
                        builder.AppendLine("");
                    }
                    else
                    {
                        // Iterate over the child DocumentObjects in the Paragraph.
                        foreach (DocumentObject obj in paragraph.ChildObjects)
                        {
                            // Check if the child DocumentObject is a TextRange.
                            if (obj.DocumentObjectType.Equals(DocumentObjectType.TextRange))
                            {
                                TextRange textRange = obj as TextRange;
                                {
                                    // Check if the TextRange is a deleted revision.
                                    if (textRange.IsDeleteRevision)
                                    {
                                        // Append information about the deleted revision to the StringBuilder.
                                        builder.AppendLine(string.Format("The section {0} paragraph {1} textrange {2} has been changed (deleted).", document.GetIndex(section), section.GetIndex(paragraph), paragraph.GetIndex(textRange)));
                                        builder.AppendLine("Author: " + textRange.DeleteRevision.Author);
                                        builder.AppendLine("DateTime: " + textRange.DeleteRevision.DateTime);
                                        builder.AppendLine("Type: " + textRange.DeleteRevision.Type);
                                        builder.AppendLine("Change Text: " + textRange.Text);
                                        builder.AppendLine("");
                                    }
                                    // Check if the TextRange is an inserted revision.
                                    else if (textRange.IsInsertRevision)
                                    {
                                        // Append information about the inserted revision to the StringBuilder.
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
            }

            // Specify the file name for the resulting text file.
            string output = "GetParagraphRevisionsDetails.txt";

            // Write the content of the StringBuilder to a text file.
            File.WriteAllText(output, builder.ToString());

            // Dispose the Document object.
            document.Dispose();

           //Launch the file
           TxtViewer(output);
            
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
