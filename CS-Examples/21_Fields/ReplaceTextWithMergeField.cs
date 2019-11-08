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

namespace ReplaceTextWithMergeField
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

            //Find the text that will be replaced
            TextSelection ts =  document.FindString("Test",true, true);

            TextRange tr = ts.GetAsOneRange();

            //Get the paragraph
            Paragraph par = tr.OwnerParagraph;

            //Get the index of the text in the paragraph
            int index = par.ChildObjects.IndexOf(tr);

            //Create a new field
            MergeField field = new MergeField(document);
            field.FieldName = "MergeField";

            //Insert field at specific position
            par.ChildObjects.Insert(index, field);

            //Remove the text
            par.ChildObjects.Remove(tr);           
          
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
