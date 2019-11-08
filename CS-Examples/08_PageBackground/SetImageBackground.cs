using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetImageBackground
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //load a word document
            Document document = new Document(@"..\..\..\..\..\..\Data\Template.docx");

            //set the background type as picture.
            document.Background.Type = BackgroundType.Picture;

            //set the background picture
            document.Background.Picture = Image.FromFile(@"..\..\..\..\..\..\Data\Background.png");

            //save the file.
            document.SaveToFile("ImageBackground.docx", FileFormat.Docx);

            //launching the Word file.
            WordDocViewer("ImageBackground.docx");


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
