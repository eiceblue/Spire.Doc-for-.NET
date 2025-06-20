using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;

namespace ToPostScript
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create word document
			Document doc = new Document();

			//Load the file from disk.
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\ConvertedTemplate.docx");

			string result = "ToPostScript.ps";

			//Save to PS file
			doc.SaveToFile(result, FileFormat.PostScript);

			//Dispose the document
			doc.Dispose();
			
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
