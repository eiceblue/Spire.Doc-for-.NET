using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ToSVG
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
			Document document = new Document();

			//Load the file from disk.
			document.LoadFromFile(@"..\..\..\..\..\..\Data\ToSVGTemplate.docx");

			//Save to SVG file.
			document.SaveToFile("Sample.svg", FileFormat.SVG);

			//Dispose the document
			document.Dispose();
			
			//Launching the svg file.
            System.Diagnostics.Process.Start("Sample.svg");
        }
    }
}
