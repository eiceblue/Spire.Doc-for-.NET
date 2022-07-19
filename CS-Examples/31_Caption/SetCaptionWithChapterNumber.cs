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
            //Create a new document
            Document document = new Document();
            //Load file from disk
            document.LoadFromFile(@"..\..\..\..\..\..\Data\SetCaptionWithChapterNumber.docx");
            //Get the first section
            Section section = document.Sections[0];
            //Label name
            string name = "Caption ";
            for (int i = 0; i < section.Body.Paragraphs.Count; i++)
            {
                for (int j = 0; j < section.Body.Paragraphs[i].ChildObjects.Count; j++)
                {
                    if (section.Body.Paragraphs[i].ChildObjects[j] is DocPicture)
                    {
                        DocPicture pic1 = section.Body.Paragraphs[i].ChildObjects[j] as DocPicture;
                        Body body = pic1.OwnerParagraph.Owner as Body;
                        if (body != null)
                        {
                            int imageIndex = body.ChildObjects.IndexOf(pic1.OwnerParagraph);
                            //Create a new paragraph
                            Paragraph para = new Paragraph(document);
                            //Set label
                            para.AppendText(name);

                            //Add caption
                            Field field1 = para.AppendField("test", FieldType.FieldStyleRef);
                            //Chapter number
                            field1.Code = " STYLEREF 1 \\s ";
                            //Chapter delimiter
                            para.AppendText(" - ");

                            //Add picture sequence number
                            SequenceField field2 = (SequenceField)para.AppendField(name, FieldType.FieldSequence);
                            field2.CaptionName = name;
                            field2.NumberFormat = CaptionNumberingFormat.Number;
                            body.Paragraphs.Insert(imageIndex + 1, para);
                        }
                    }
                }
            }
            //Set update fields
            document.IsUpdateFields = true;
            //Save the result file
            string output = "SetCaptionWithChapterNumber.docx";
            document.SaveToFile(output, FileFormat.Docx);

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
