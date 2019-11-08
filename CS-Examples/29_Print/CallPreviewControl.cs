using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing.Printing;

namespace CallPreviewControl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load the document
            string input = @"..\..\..\..\..\..\Data\Sample.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the PrintDocument object
            PrintDocument printDoc = doc.PrintDocument;

            //Call print preview dialog
            PrintPreviewDialog printPreviewDialog = new PrintPreviewDialog();
            printPreviewDialog.Document = doc.PrintDocument;

            //Set the preview dialog size of client area
            printPreviewDialog.ClientSize = new Size(600, 800);
            printPreviewDialog.ShowDialog();
        }
    }
}
