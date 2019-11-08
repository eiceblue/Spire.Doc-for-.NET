using System;
using System.Windows.Forms;
using Spire.Doc;
using System.Drawing.Printing;

namespace CustomPaperSize
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

            //Custom the paper size
            printDoc.DefaultPageSettings.PaperSize = new PaperSize("custom", 900, 800);

            //Print the document
            printDoc.Print();
        }
    }
}
