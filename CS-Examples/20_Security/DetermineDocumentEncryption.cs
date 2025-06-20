using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace DetermineDocumentEncryption
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool isEncrypted = Document.IsEncrypted(@"..\..\..\..\..\..\Data\TemplateWithPassword.docx");
            if(isEncrypted == true)
            {
                MessageBox.Show("This document is encrypted. ");
            }
            else
            {
                MessageBox.Show("This document is unencrypted. ");
            }
        }

    }
}
