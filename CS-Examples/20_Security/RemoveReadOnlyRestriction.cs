using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;

namespace RemoveReadOnlyRestriction
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

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveReadOnlyRestriction.docx");

            //Remove ReadOnly Restriction.
            document.Protect(ProtectionType.NoProtection);

            String result = "RemoveReadOnlyRestriction_out.docx";

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
