using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetTransparencyForTextbox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a word document
			Document doc = new Document();

			//Create a new section
			Section section = doc.AddSection();

			//Create a new paragraph
			Paragraph paragraph = section.AddParagraph();

			//Append TextBox
			Spire.Doc.Fields.TextBox textbox1 = paragraph.AppendTextBox(100, 50);

			//Set fill color
			textbox1.Format.FillColor = Color.Red;

			//Set fill transparency
			textbox1.FillTransparency = 0.45;

			//Save the Word file
			string output = "SetTransparencyForTextbox.docx";
			doc.SaveToFile(output, FileFormat.Docx2013);

			// Dispose the document
			doc.Dispose();

            //Launch the file 
            WordDocViewer(output);
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
