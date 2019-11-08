using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddVariables
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create Word document.
            Document document = new Document();

            //Add a section.
            Section section = document.AddSection();

            //Add a paragraph.
            Paragraph paragraph = section.AddParagraph();

            //Add a DocVariable field.
            paragraph.AppendField("A1", FieldType.FieldDocVariable);

            //Add a document variable to the DocVariable field.
            document.Variables.Add("A1", "12");

            //Update fields.
            document.IsUpdateFields = true;

            String result = "Result-AddVariables.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the MS Word file.
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
