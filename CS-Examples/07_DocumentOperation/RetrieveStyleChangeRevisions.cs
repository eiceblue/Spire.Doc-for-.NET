using System;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Fields;
using System.IO;

namespace RetrieveStyleChangeRevisions
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize a new, empty Document object.
            Document doc = new Document();

            // Load the existing Word document containing revision history from the specified file path.
            doc.LoadFromFile(@"..\..\..\..\..\..\..\Data\GetRevisions.docx");

            // Retrieve the collection of all revision information (changes, comments, etc.) from the document.
            RevisionInfoCollection revisionInfoCollection = doc.GetRevisionInfos();

            // Initialize a StringBuilder to efficiently construct the output text report.
            StringBuilder stringBuilder = new StringBuilder();

            // Iterate through each revision item in the collected revision information.
            foreach (RevisionInfo revisionInfo in revisionInfoCollection)
            {
                // Check if the current revision is specifically a formatting change (e.g., bold, color, font).
                if (revisionInfo.RevisionType == RevisionType.FormatChange)
                {
                    // Verify if the object affected by this revision is a TextRange (a segment of text).
                    if (revisionInfo.OwnerObject is Spire.Doc.Fields.TextRange)
                    {
                        // Cast the owner object to a TextRange to access its specific properties.
                        TextRange range = (TextRange)revisionInfo.OwnerObject;

                        // Append the actual text content of the modified range to the report.
                        stringBuilder.AppendLine("TextRange:" + range.Text + "rn");

                        // Switch the document view to the "Original" state to read pre-change formatting properties.
                        doc.RevisionsView = RevisionsView.Original;

                        // Append the original formatting details (Bold, Color, Highlight, Font, Underline) to the report.
                        stringBuilder.AppendLine("Original style£º" + "isBold£º" + range.CharacterFormat.Bold + ";" + "TextColor£º" + range.CharacterFormat.TextColor + "£»HighlightColor£º" + range.CharacterFormat.HighlightColor + "£»FontName£º" + range.CharacterFormat.FontName + "£»UnderlineStyle£º" + range.CharacterFormat.UnderlineStyle + "rn");

                        // Switch the document view to the "Final" state to read post-change formatting properties.
                        doc.RevisionsView = RevisionsView.Final;

                        // Append the final formatting details to compare against the original state.
                        stringBuilder.AppendLine("Final style£º" + "isBold£º" + range.CharacterFormat.Bold + ";" + "TextColor£º" + range.CharacterFormat.TextColor + "£»HighlightColor£º" + range.CharacterFormat.HighlightColor + "£»FontName£º" + range.CharacterFormat.FontName + "£»UnderlineStyle£º" + range.CharacterFormat.UnderlineStyle + "rn");
                    }
                }
            }

            // Write the complete accumulated report string to a text file.
            File.WriteAllText("RetrieveStyleChangeRevisions.txt", stringBuilder.ToString());

            // Close the document to release file resources and memory.
            doc.Close();
        }
    }
}
