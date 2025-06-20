using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Collections.Generic;

namespace ReplaceWithHtml
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Read HTML content from a file and store it in a String variable
            String HTML = File.ReadAllText(@"..\..\..\..\..\..\Data\InputHtml1.txt");

            // Create a new Document object from a template file and store it in a Document variable
            Document document = new Document(@"..\..\..\..\..\..\Data\ReplaceWithHtml.docx");

            // Create a new List to store DocumentObject objects
            List<DocumentObject> replacement = new List<DocumentObject>();

            // Add a new section to the document and store it in a Section variable
            Section tempSection = document.AddSection();

            // Add a new paragraph to the section and store it in a Paragraph variable
            Paragraph par = tempSection.AddParagraph();
            // Append the HTML content to the paragraph
            par.AppendHTML(HTML);

            // Iterate through all child objects of the section
            foreach (DocumentObject obj in tempSection.Body.ChildObjects)
            {
                // Cast the current object to DocumentObject and add it to the replacement list
                DocumentObject docObj = obj as DocumentObject;
                replacement.Add(docObj);
            }

            // Find all occurrences of the string "[#placeholder]" in the document and store them in an array of TextSelection objects
            TextSelection[] selections = document.FindAllString("[#placeholder]", false, true);

            // Create a new List to store TextRangeLocation objects
            List<TextRangeLocation> locations = new List<TextRangeLocation>();
            // Iterate through all TextSelection objects found
            foreach (TextSelection selection in selections)
            {
                // Get the range of the current selection and create a new TextRangeLocation object with it
                locations.Add(new TextRangeLocation(selection.GetAsOneRange()));
            }
            // Sort the locations list in ascending order
            locations.Sort();

            // Iterate through all TextRangeLocation objects in the locations list
            foreach (TextRangeLocation location in locations)
            {
                // Call the ReplaceWithHTML method with the current location and replacement list as arguments
                ReplaceWithHTML(location, replacement);
            }

            // Remove the temporary section added earlier from the document
            document.Sections.Remove(tempSection);

            // Save the modified document to a file named "Output.docx" with the Docx file format
            document.SaveToFile("Output.docx", FileFormat.Docx);
            // Dispose of the document object to release resources
            document.Dispose();

            //Launch the Word file.
            WordDocViewer("Output.docx");
        }

        private static void ReplaceWithHTML(TextRangeLocation location, List<DocumentObject> replacement)
        {
            TextRange textRange = location.Text;

            //textRange index
            int index = location.Index;

            //get owner paragraph
            Paragraph paragraph = location.Owner;

            //get owner text body
            Body sectionBody = paragraph.OwnerTextBody;

            //get the index of paragraph in section
            int paragraphIndex = sectionBody.ChildObjects.IndexOf(paragraph);

            int replacementIndex = -1;
            if (index == 0)
            {
                //remove the first child object
                paragraph.ChildObjects.RemoveAt(0);

                replacementIndex = sectionBody.ChildObjects.IndexOf(paragraph);
            }
            else if (index == paragraph.ChildObjects.Count - 1)
            {
                paragraph.ChildObjects.RemoveAt(index);
                replacementIndex = paragraphIndex + 1;
            }
            else
            {
                //split owner paragraph
                Paragraph paragraph1 = (Paragraph)paragraph.Clone();
                while (paragraph.ChildObjects.Count > index)
                {
                    paragraph.ChildObjects.RemoveAt(index);
                }
                int i = 0;
                int count = index + 1;
                while (i < count)
                {
                    paragraph1.ChildObjects.RemoveAt(0);
                    i += 1;
                }
                sectionBody.ChildObjects.Insert(paragraphIndex + 1, paragraph1);

                replacementIndex = paragraphIndex + 1;
            }

            //insert replacement
            for (int i = 0; i <= replacement.Count - 1; i++)
            {
                sectionBody.ChildObjects.Insert(replacementIndex + i, replacement[i].Clone());
            }
        }

        public class TextRangeLocation : IComparable<TextRangeLocation>
        {
            public TextRangeLocation(TextRange text)
            {
                this.Text = text;
            }

            public TextRange Text
            {
                get { return m_Text; }
                set { m_Text = value; }
            }

            private TextRange m_Text;
            public Paragraph Owner
            {
                get { return this.Text.OwnerParagraph; }
            }

            public int Index
            {
                get { return this.Owner.ChildObjects.IndexOf(this.Text); }
            }

            public int CompareTo(TextRangeLocation other)
            {
                return -(this.Index - other.Index);
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
