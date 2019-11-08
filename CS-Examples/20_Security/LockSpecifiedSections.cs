using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace LockSpecifiedSections
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

            //Add new sections.
            Section s1 = document.AddSection();
            Section s2 = document.AddSection();
           
            //Append some text to section 1 and section 2.
            s1.AddParagraph().AppendText("Spire.Doc demo, section 1");
            s2.AddParagraph().AppendText("Spire.Doc demo, section 2");

            //Protect the document with AllowOnlyFormFields protection type.
            document.Protect(ProtectionType.AllowOnlyFormFields, "123");

            //Unprotect section 2
            s2.ProtectForm = false;

            String result = "Result-LockSpecifiedSections.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the file.
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
