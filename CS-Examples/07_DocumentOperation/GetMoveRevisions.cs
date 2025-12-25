using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Formatting.Revisions;
using Spire.Doc.Fields;
using System.IO;

namespace GetMoveRevisions
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Load the existing Word document that contains tracked moves
            Document document = new Document(@"..\..\..\..\..\..\..\Data\MoveRevision.docx");

            // Create a DifferRevisions object to access move revisions within the document
            DifferRevisions differRevisions = new DifferRevisions(document);

            // Get the list of 'Move From' revisions (content that was moved from one location)
            List<DocumentObject> moveFromRevisions = differRevisions.MoveFromRevisions;

            // Get the list of 'Move To' revisions (content that was moved to a new location)
            List<DocumentObject> moveToRevisions = differRevisions.MoveToRevisions;

            // Create a StringBuilder to accumulate information about 'Move From' revisions
            StringBuilder moveFromRevisions_content = new StringBuilder();

            // Append a header line indicating the count of 'Move From' revisions
            moveFromRevisions_content.AppendLine("MoveFromRevisions: " + moveFromRevisions.Count);

            // Loop through each 'Move From' revision object
            for (int i = 0; i < moveFromRevisions.Count; i++)
            {
                // Append the string representation of the revision object
                moveFromRevisions_content.AppendLine(moveFromRevisions[i].ToString());

                // Check if the revision object is a Paragraph
                if (moveFromRevisions[i].DocumentObjectType == DocumentObjectType.Paragraph)
                {
                    // If it's a paragraph, append its text content
                    moveFromRevisions_content.AppendLine(((Paragraph)moveFromRevisions[i]).Text);
                }

                // Check if the revision object is a TextRange (a piece of text)
                if (moveFromRevisions[i].DocumentObjectType == DocumentObjectType.TextRange)
                {
                    // If it's a text range, append its text content
                    moveFromRevisions_content.AppendLine(((TextRange)moveFromRevisions[i]).Text);
                }
            }

            // Create a StringBuilder to accumulate information about 'Move To' revisions
            StringBuilder moveToRevisions_content = new StringBuilder();

            // Append a header line indicating the count of 'Move To' revisions
            moveToRevisions_content.AppendLine("MoveToRevisions: " + moveToRevisions.Count);

            // Loop through each 'Move To' revision object
            for (int i = 0; i < moveToRevisions.Count; i++)
            {
                // Append the string representation of the revision object
                moveToRevisions_content.AppendLine(moveToRevisions[i].ToString());

                // Check if the revision object is a Paragraph
                if (moveToRevisions[i].DocumentObjectType == DocumentObjectType.Paragraph)
                {
                    // If it's a paragraph, append its text content
                    moveToRevisions_content.AppendLine(((Paragraph)moveToRevisions[i]).Text);
                }

                // Check if the revision object is a TextRange (a piece of text)
                if (moveToRevisions[i].DocumentObjectType == DocumentObjectType.TextRange)
                {
                    // If it's a text range, append its text content
                    moveToRevisions_content.AppendLine(((TextRange)moveToRevisions[i]).Text);
                }
            }

            // Write the accumulated 'Move From' revision information to a text file
            File.WriteAllText("MoveFromRevisions.txt", moveFromRevisions_content.ToString());

            // Write the accumulated 'Move To' revision information to a text file
            File.WriteAllText("MoveToRevisions.txt", moveToRevisions_content.ToString());

            // Dispose of the document object to free resources
            document.Dispose();
        }
    }
}
