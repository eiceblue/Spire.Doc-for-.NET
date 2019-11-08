using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing.Printing;

namespace SetMarginAndDuplex
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

            //Set graphics origin starts at the page margins
            printDoc.OriginAtMargins = true;
            //Set the margin to 0
            printDoc.DefaultPageSettings.Margins = new System.Drawing.Printing.Margins(0, 0, 0, 0);

            //Double-sided, vertical printing
            printDoc.PrinterSettings.Duplex = Duplex.Vertical;
            //Double-sided, horizontal printing
            //printDoc.PrinterSettings.Duplex = Duplex.Horizontal;

            //Print the word document
            printDoc.Print();
        }
    }
}
