using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;
using Spire.Doc.Fields;

namespace ModifyHyperlinkText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load Document
            string input = @"..\..\..\..\..\..\Data\Hyperlinks.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Find all hyperlinks in the Word document
            List<Field> hyperlinks = new List<Field>();
            foreach (Section section in doc.Sections)
            {
                foreach (DocumentObject sec in section.Body.ChildObjects)
                {
                    if (sec.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
                        {
                            if (para.DocumentObjectType == DocumentObjectType.Field)
                            {
                                Field field = para as Field;

                                if (field.Type == FieldType.FieldHyperlink)
                                {
                                    hyperlinks.Add(field);
                                }
                            }
                        }
                    }
                }
            }

            //Reset the property of hyperlinks[0].FieldText by using the index of the hyperlink
            hyperlinks[0].FieldText = "Spire.Doc component";

            //Save and launch document
            string output = "ModifyText.docx";
            doc.SaveToFile(output, FileFormat.Docx);
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
