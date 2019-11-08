using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

namespace AddRichTextContentControl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a document
            Document document = new Document();

            //Add a new section.
            Section section = document.AddSection();

            //Add a paragraph
            Paragraph paragraph = section.AddParagraph();

            //Append textRange for the paragraph
            TextRange txtRange = paragraph.AppendText("The following example shows how to add RichText content control in a Word document. \n");

            //Append textRange 
            txtRange = paragraph.AppendText("Add RichText Content Control:  ");

            //Set the font format
            txtRange.CharacterFormat.Italic = true;

            //Create StructureDocumentTagInline for document
            StructureDocumentTagInline sdt = new StructureDocumentTagInline(document);

            //Add sdt in paragraph
            paragraph.ChildObjects.Add(sdt);

            //Specify the type
            sdt.SDTProperties.SDTType = SdtType.RichText;

            //Set displaying text
            SdtText text = new SdtText(true);
            text.IsMultiline = true;
            sdt.SDTProperties.ControlProperties = text;

            //Crate a TextRange
            TextRange rt = new TextRange(document);
            rt.Text = "Welcome to use ";
            rt.CharacterFormat.TextColor = Color.Green;
            sdt.SDTContent.ChildObjects.Add(rt);

            rt = new TextRange(document);
            rt.Text = "Spire.Doc";
            rt.CharacterFormat.TextColor = Color.OrangeRed;
            sdt.SDTContent.ChildObjects.Add(rt);

            //Save the document.
            document.SaveToFile("Output.docx", FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer("Output.docx");
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
