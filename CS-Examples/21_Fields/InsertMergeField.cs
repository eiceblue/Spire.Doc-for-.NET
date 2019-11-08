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

namespace InsertMergeField
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

            //Add merge field in the paragraph
            MergeField field
            = par.AppendField("MyFieldName", FieldType.FieldMergeField) as MergeField;
           
            //Save to file
            document.SaveToFile("result.docx", FileFormat.Docx);
            //Launch result file
            WordDocViewer("result.docx");

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
