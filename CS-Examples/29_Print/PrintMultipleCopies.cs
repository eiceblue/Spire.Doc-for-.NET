using Spire.Doc;
using System;
using System.Windows.Forms;

namespace PrintMultipleCopies
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load word document
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

            document.PrintDocument.PrinterSettings.PrinterName = "Microsoft Print to PDF";
            document.PrintDocument.PrinterSettings.Copies = 4;

            document.PrintDocument.Print();

        }

    }
}
