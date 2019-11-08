using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetGradientBackground
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_2.docx");

            //Set the background type as Gradient.
            document.Background.Type = BackgroundType.Gradient;
            BackgroundGradient Test = document.Background.Gradient;

            //Set the first color and second color for Gradient.
            Test.Color1 = Color.White;
            Test.Color2 = Color.LightBlue;

            //Set the Shading style and Variant for the gradient.
            Test.ShadingVariant = GradientShadingVariant.ShadingDown;
            Test.ShadingStyle = GradientShadingStyle.Horizontal;

            String result = "Result-SetGradientBackground.docx";

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
