using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace DetectAndRemoveVBAMacros
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\DetectAndRemoveVBAMacros.docm");

            //If the document contains Macros, remove them from the document.
            if (document.IsContainMacro)
            {
                document.ClearMacros();
            }

            String result = "Result-DetectAndRemoveVBAMacros.docm";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docm);

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
