using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;

namespace GetTextByStyleName
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a Word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\Template_N5.docx");

			//Create string builder
			StringBuilder builder = new StringBuilder();

			//Loop through sections
			foreach (Section section in doc.Sections)
			{
				//Loop through paragraphs
				foreach (Paragraph para in section.Paragraphs)
				{
					//Find the paragraph whose style name is "Heading1"
					if (para.StyleName == "Heading1")
					{
						//Write the text of paragraph
						builder.AppendLine(para.Text);
					}
				}
			}

			//Write the contents in a TXT file
			string output = "GetTextByStyleName_out.txt";
			File.WriteAllText(output, builder.ToString());

			//Dispose the document
			doc.Dispose();

            //Launch the file
            TxtViewer(output);
        }
        private void TxtViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}
