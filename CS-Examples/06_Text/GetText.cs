using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using System.IO;

namespace GetText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {   
			//Load the document from disk.
            Document document = new Document(@"..\..\..\..\..\..\Data\ExtractText.docx");

            //get text from document
            string text = document.GetText();

            //create a new TXT File to save extracted text
            File.WriteAllText("Extract.txt", text);

            //launch the file.
            WordDocViewer("Extract.txt");
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
