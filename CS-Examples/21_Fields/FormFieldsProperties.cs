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

namespace FormFieldsProperties
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
            Document document = new Document(@"..\..\..\..\..\..\Data\FillFormField.doc");

            //Get the first section
            Section section = document.Sections[0];

            //Get FormField by index
            FormField formField = section.Body.FormFields[1];

            if (formField.Type == FieldType.FieldFormTextInput)
            {
                formField.Text = "My name is " + formField.Name;
                formField.CharacterFormat.TextColor = Color.Red;
                formField.CharacterFormat.Italic = true;
            }


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
