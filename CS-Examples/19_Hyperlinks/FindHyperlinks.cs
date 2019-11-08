using System;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Collections.Generic;

namespace FindHyperlinks
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

            //Create a hyperlink list
            List<Field> hyperlinks = new List<Field>();
            string hyperlinksText = null;
            //Iterate through the items in the sections to find all hyperlinks
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
                                    //Get the hyperlink text
                                    hyperlinksText += field.FieldText + "\r\n";
                                }
                            }
                        }
                    }
                }
            }

            //Save the text of all hyperlinks to TXT File and launch it
            string output = "HyperlinksText.txt";
            File.WriteAllText(output, hyperlinksText);
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
