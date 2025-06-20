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
            //Create a Word document
			Document document = new Document();

			//Load the file from disk
			document.LoadFromFile(@"..\..\..\..\..\..\Data\MailMerge.doc");

			//prepare sample data
			string[] fieldNames = new string[] { "Contact Name", "Fax", "Date" };

			string[] fieldValues = new string[] { "John Smith", "+1 (69) 123456", System.DateTime.Now.Date.ToString() };

			//Begin the mail merge process
			document.MailMerge.Execute(fieldNames, fieldValues);

			//Save the document.
			document.SaveToFile("Sample.doc", FileFormat.Doc);

			// Dispose the document object
			document.Dispose();

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
