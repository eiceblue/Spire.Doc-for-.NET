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

namespace InsertPageRefField
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
            Document document = new Document(@"..\..\..\..\..\..\Data\PageRef.docx");

            //Get the first section
            Section section = document.LastSection;

            Paragraph par = section.AddParagraph();

            //Add page ref field
            Field field
            = par.AppendField("pageRef", FieldType.FieldPageRef);

            //Set field code
            field.Code = "PAGEREF  bookmark1 \\# \"0\" \\* Arabic  \\* MERGEFORMAT";

            //Update field
            document.IsUpdateFields = true;
           
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
