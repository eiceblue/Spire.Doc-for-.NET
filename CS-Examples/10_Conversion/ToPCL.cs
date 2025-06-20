using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;

namespace ToPCL
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

            		//On Net4.6 and above platforms with adding the following external dependencies, you can set the UseHarfBuzzTextShaper which can better handling Thai and Tibetan characters
            		//external reference to:  
            		//HarfBuzzSharp >= 2.6.1.5
            		//System.Buffers >= 4.4.0
            		//System.Memory >= 4.5.3
            		//System.Numerics.Vectors >= 4.4.0
            		//System.Runtime.CompilerServices.Unsafe >= 4.5.2

            		//document.LayoutOptions.UseHarfBuzzTextShaper = true;

			string result = "ToPCL.pcl";

			//Save to PCL file
			doc.SaveToFile(result, FileFormat.PCL);

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
