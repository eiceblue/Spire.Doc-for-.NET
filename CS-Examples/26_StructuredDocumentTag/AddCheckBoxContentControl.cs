using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AddCheckBoxContentControl
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
            TextRange txtRange = paragraph.AppendText("The following example shows how to add CheckBox content control in a Word document. \n");

            // Append text indicating adding the CheckBox content control
            txtRange = paragraph.AppendText("Add CheckBox Content Control:  ");

            // Set the text range formatting to italic
            txtRange.CharacterFormat.Italic = true;

            // Create an inline structure document tag (SDT) and add it to the paragraph's child objects
            StructureDocumentTagInline sdt = new StructureDocumentTagInline(document);
            paragraph.ChildObjects.Add(sdt);

            // Set the SDT type to CheckBox
            sdt.SDTProperties.SDTType = SdtType.CheckBox;

            // Create an instance of SdtCheckBox and set it as the control properties for the SDT
            SdtCheckBox scb = new SdtCheckBox();
            sdt.SDTProperties.ControlProperties = scb;

            // Create a TextRange object, set its font name and size
            TextRange tr = new TextRange(document);
            tr.CharacterFormat.FontName = "MS Gothic";
            tr.CharacterFormat.FontSize = 12;

            // Add the TextRange object to the SDT's child objects
            sdt.ChildObjects.Add(tr);

            // Set the CheckBox as checked
            scb.Checked = true;

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
