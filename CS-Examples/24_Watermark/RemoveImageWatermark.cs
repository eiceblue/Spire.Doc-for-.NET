using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemoveImageWatermark
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveImageWatermark.docx");

            //Set the watermark as null to remove the text and image watermark.
            document.Watermark = null;

            String result = "Result-RemoveImageWatermark.docx";

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
