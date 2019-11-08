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
            //Create a document
            Document document = new Document();

            //Add a new section.
            Section section = document.AddSection();

            //Add a paragraph
            Paragraph paragraph = section.AddParagraph();

            //Append textRange for the paragraph
            TextRange txtRange = paragraph.AppendText("The following example shows how to add CheckBox content control in a Word document. \n");

            //Append textRange 
            txtRange = paragraph.AppendText("Add CheckBox Content Control:  ");

            //Set the font format
            txtRange.CharacterFormat.Italic = true;

            //Create StructureDocumentTagInline for document
            StructureDocumentTagInline sdt = new StructureDocumentTagInline(document);

            //Add sdt in paragraph
            paragraph.ChildObjects.Add(sdt);

            //Specify the type
            sdt.SDTProperties.SDTType = SdtType.CheckBox;

            //Set properties for control
            SdtCheckBox scb = new SdtCheckBox();
            sdt.SDTProperties.ControlProperties = scb;

            //Add textRange format
            TextRange tr = new TextRange(document);
            tr.CharacterFormat.FontName = "MS Gothic";
            tr.CharacterFormat.FontSize = 12;

            //Add textRange to StructureDocumentTagInline
            sdt.ChildObjects.Add(tr);

            //Set checkBox as checked
            scb.Checked = true;

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
