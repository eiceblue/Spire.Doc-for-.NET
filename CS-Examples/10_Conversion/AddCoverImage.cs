using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Fields;
namespace AddCoverImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\..\Data\ToEpub.doc");
            DocPicture picture = new DocPicture(doc);
            picture.LoadImage(Image.FromFile(@"..\..\..\..\..\..\..\Data\Cover.png"));
            string result = "AddCoverImage.epub";
            doc.SaveToEpub(result, picture);
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
