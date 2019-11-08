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
            String HTML = File.ReadAllText(@"..\..\..\..\..\..\Data\InputHtml1.txt");

            //Load the document from disk.  
            Document document = new Document(@"..\..\..\..\..\..\Data\ReplaceWithHtml.docx");

            //collect the objects which is used to replace text
            List<DocumentObject> replacement = new List<DocumentObject>();

            //create a temporary section
            Section tempSection = document.AddSection();

            //add a paragraph to append html
            Paragraph par = tempSection.AddParagraph();
            par.AppendHTML(HTML);

            //get the objects in temporary section
            foreach (DocumentObject obj in tempSection.Body.ChildObjects)
            {
                DocumentObject docObj = obj as DocumentObject;
                replacement.Add(docObj);
            }

            //Find all text which will be replaced.
            TextSelection[] selections = document.FindAllString("[#placeholder]", false, true);

            List<TextRangeLocation> locations = new List<TextRangeLocation>();
            foreach (TextSelection selection in selections)
            {
                locations.Add(new TextRangeLocation(selection.GetAsOneRange()));
            }
            locations.Sort();

            foreach (TextRangeLocation location in locations)
            {
                //replace the text with HTML
                ReplaceWithHTML(location, replacement);
            }

            //remove the temp section
            document.Sections.Remove(tempSection);

            //Save the document.
            document.SaveToFile("Output.docx", FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer("Output.docx");
        }

        private static void ReplaceWithHTML(TextRangeLocation location, List<DocumentObject> replacement)
        {
            TextRange textRange = location.Text;

            //textRange index
            int index = location.Index;

            //get owener paragraph
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
