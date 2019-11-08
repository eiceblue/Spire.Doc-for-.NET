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

namespace GetFormFieldsCollection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            //Open a Word document
            Document document = new Document(@"..\..\..\..\..\..\Data\FillFormField.doc");

            //Get the first section
            Section section = document.Sections[0];

            FormFieldCollection formFields = section.Body.FormFields;

            sb.Append("The first section has " + formFields.Count + " form fields.");

            File.WriteAllText("result.txt", sb.ToString());

            //Launch result file
            WordDocViewer("result.txt");

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
