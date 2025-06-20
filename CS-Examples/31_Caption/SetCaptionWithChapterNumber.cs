using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetCaptionWithChapterNumber
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of Document
            Document document = new Document();

            // Load the Word document from the specified file
            document.LoadFromFile(@"..\..\..\..\..\..\Data\SetCaptionWithChapterNumber.docx");

            // Get the first section of the document
            Section section = document.Sections[0];

            // Specify the base name for the captions
            string name = "Caption ";

            // Iterate through paragraphs in the body of the section
            for (int i = 0; i < section.Body.Paragraphs.Count; i++)
            {
                // Iterate through child objects within each paragraph
                for (int j = 0; j < section.Body.Paragraphs[i].ChildObjects.Count; j++)
                {
                    // Check if the child object is a picture
                    if (section.Body.Paragraphs[i].ChildObjects[j] is DocPicture)
                    {
                        // Convert the child object to a DocPicture
                        DocPicture pic1 = section.Body.Paragraphs[i].ChildObjects[j] as DocPicture;

                        // Get the owner paragraph's owner, which should be the Body
                        Body body = pic1.OwnerParagraph.Owner as Body;

                        if (body != null)
                        {
                            // Find the index of the owner paragraph within the Body
                            int imageIndex = body.ChildObjects.IndexOf(pic1.OwnerParagraph);

                            // Create a new paragraph
                            Paragraph para = new Paragraph(document);

                            // Append the caption name
                            para.AppendText(name);

                            // Append a field for referencing the chapter number using a style reference
                            Field field1 = para.AppendField("test", FieldType.FieldStyleRef);
                            field1.Code = " STYLEREF 1 \\s ";

                            // Append a separator text
                            para.AppendText(" - ");

                            // Append a sequence field for the caption number
                            SequenceField field2 = (SequenceField)para.AppendField(name, FieldType.FieldSequence);
                            field2.CaptionName = name;
                            field2.NumberFormat = CaptionNumberingFormat.Number;

                            // Insert the new paragraph after the owner paragraph
                            body.Paragraphs.Insert(imageIndex + 1, para);
                        }
                    }
                }
            }

            // Enable field updating in the document
            document.IsUpdateFields = true;

            // Specify the output file name and format (Docx)
            string output = "SetCaptionWithChapterNumber.docx";
            document.SaveToFile(output, FileFormat.Docx);

            // Dispose of the document object when finished using it
            document.Dispose();

            //Launching the file
            WordDocViewer(output);

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
