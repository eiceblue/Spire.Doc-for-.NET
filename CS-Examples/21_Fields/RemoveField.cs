using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Collections;

namespace RemoveField
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
            Document document = new Document(@"..\..\..\..\..\..\Data\IfFieldSample.docx");

            //Get the first field
            Field field = document.Fields[0];

            //Get the paragraph of the field
            Paragraph par = field.OwnerParagraph;
            //Get the index of the  field
            int index = par.ChildObjects.IndexOf(field);
            //Remove if field via index
            par.ChildObjects.RemoveAt(index);

            //Save doc file
            document.SaveToFile("result.docx",FileFormat.Docx);

            //Launch the Word file
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
