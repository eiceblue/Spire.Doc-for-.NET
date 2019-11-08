using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Collections;
using System.Text;

namespace SetLocaleForField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Open a Word document
            Document document = new Document(@"..\..\..\..\..\..\Data\SampleB_2.docx");

            //Get the first section
            Section section = document.Sections[0];

            Paragraph par = section.AddParagraph();

            //Add a date field
            Field field
            = par.AppendField("DocDate", FieldType.FieldDate);

            //Set the LocaleId for the textrange
            (field.OwnerParagraph.ChildObjects[0] as TextRange).CharacterFormat.LocaleIdASCII = 1049;

            field.FieldText = "2019-10-10";
            //Update field
            document.IsUpdateFields = true;

            String result= "result.docx";
            document.SaveToFile(result, FileFormat.Docx);
            //Launch result file
            WordDocViewer(result);

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
