using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Collections;

namespace ConvertIfFieldToText
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


            //Get all fields in document
            FieldCollection fields = document.Fields;

            for (int i = 0; i < fields.Count; i++)
            {
                Field field = fields[i];
                if (field.Type == FieldType.FieldIf)
                {
                    TextRange original = field as TextRange;
                    //Get field text
                    string text = field.FieldText;
                    //Create a new textRange and set its format
                    TextRange textRange = new TextRange(document);
                    textRange.Text = text;
                    textRange.CharacterFormat.FontName = original.CharacterFormat.FontName;
                    textRange.CharacterFormat.FontSize = original.CharacterFormat.FontSize;

                    Paragraph par = field.OwnerParagraph;
                    //Get the index of the if field
                    int index = par.ChildObjects.IndexOf(field);
                    //Remove if field via index
                    par.ChildObjects.RemoveAt(index);
                    //Insert field text at the position of if field
                    par.ChildObjects.Insert(index, textRange);
                }

            }

            String result ="result.docx";
            //Save doc file
            document.SaveToFile(result, FileFormat.Docx);

            //Launch the Word file
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
