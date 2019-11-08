using System;
using System.Windows.Forms;
using Spire.Doc;

namespace MailMerage
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\MailMerge.doc");

            string[] filedNames = new string[]{"Contact Name","Fax","Date"};

            string[] filedValues = new string[]{"John Smith","+1 (69) 123456",System.DateTime.Now.Date.ToString()};

            document.MailMerge.Execute(filedNames, filedValues);

          
            //Save doc file.
            document.SaveToFile("Sample.doc", FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer("Sample.doc");
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
