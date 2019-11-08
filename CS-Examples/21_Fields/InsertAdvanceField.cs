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

namespace InsertAdvanceField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Open a Word document.
            Document document = new Document(@"..\..\..\..\..\..\Data\SampleB_2.docx");

            //Get the first section
            Section section = document.Sections[0];

            Paragraph par = section.AddParagraph();
           
            //Add advance field
            Field field
            = par.AppendField("Field", FieldType.FieldAdvance);

            //Add field code
            field.Code = "ADVANCE \\d 10 \\l 10 \\r 10 \\u 0 \\x 100 \\y 100 ";

            //Update field
            document.IsUpdateFields = true;

            String result="result.docx";
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
