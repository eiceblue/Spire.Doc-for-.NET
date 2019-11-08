using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetCultureForField
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

            //Create a section
            Section section = document.AddSection();

            //Add paragraph
            Paragraph paragraph = section.AddParagraph();

            //Add textRnage for paragraph
            paragraph.AppendText("Add Date Field: ");

            //Add date field1
            Field field1 = paragraph.AppendField("Date1", FieldType.FieldDate) as Field;
            field1.Code = @"DATE  \@" + "\"yyyy\\MM\\dd\"";

            //Add new paragraph
            Paragraph newParagraph = section.AddParagraph();

            //Add textRnage for paragraph
            newParagraph.AppendText("Add Date Field with setting French Culture: ");

            //Add date field2
            Field field2 = newParagraph.AppendField("\"\\@\"dd MMMM yyyy", FieldType.FieldDate);

            //Setting Field with setting French Culture
            field2.CharacterFormat.LocaleIdASCII = 1036;

            //Update fields
            document.IsUpdateFields = true;

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
