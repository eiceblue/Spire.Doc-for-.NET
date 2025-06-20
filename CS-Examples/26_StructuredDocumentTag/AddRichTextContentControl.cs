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
      
            // Create a new document object
            Document document = new Document();

            // Add a section to the document
            Section section = document.AddSection();

            // Add a paragraph to the section
            Paragraph paragraph = section.AddParagraph();

            // Append text with an explanation to the paragraph
            TextRange txtRange = paragraph.AppendText("The following example shows how to add RichText content control in a Word document. \n");

            // Append text indicating adding the RichText content control
            txtRange = paragraph.AppendText("Add RichText Content Control:  ");

            // Set the text range formatting to italic
            txtRange.CharacterFormat.Italic = true;

            // Create an inline structure document tag (SDT) and add it to the paragraph's child objects
            StructureDocumentTagInline sdt = new StructureDocumentTagInline(document);
            paragraph.ChildObjects.Add(sdt);

            // Set the SDT type to RichText
            sdt.SDTProperties.SDTType = SdtType.RichText;

            // Create an instance of SdtText, set its multiline property, and assign it as the control properties for the SDT
            SdtText text = new SdtText(true);
            text.IsMultiline = true;
            sdt.SDTProperties.ControlProperties = text;

            // Create a TextRange object and set its text and text color, then add it to the SDT's content
            TextRange rt = new TextRange(document);
            rt.Text = "Welcome to use ";
            rt.CharacterFormat.TextColor = Color.Green;
            sdt.SDTContent.ChildObjects.Add(rt);

            // Create another TextRange object and set its text and text color, then add it to the SDT's content
            rt = new TextRange(document);
            rt.Text = "Spire.Doc";
            rt.CharacterFormat.TextColor = Color.OrangeRed;
            sdt.SDTContent.ChildObjects.Add(rt);

            // Save the document to a file in Docx format
            document.SaveToFile("Output.docx", FileFormat.Docx);

            // Dispose the document object
            document.Dispose();
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
