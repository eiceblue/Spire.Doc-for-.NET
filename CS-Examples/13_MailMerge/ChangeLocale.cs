using System;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ChangeLocale
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

			// Store the current culture so it can be set back once mail merge is complete.
			CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;

			//Set the current thread culture
			Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

			//prepare sample data 
			string[] fieldNames = new string[] { "Contact Name", "Fax", "Date" };
			string[] fieldValues = new string[] { "John Smith", "+1 (69) 123456", System.DateTime.Now.ToString() };

			//excute mail merge
			document.MailMerge.Execute(fieldNames, fieldValues);

			//restore the thread culture
			Thread.CurrentThread.CurrentCulture = currentCulture;

			//Save doc file.
			string output = "ChangeLocale.docx";
			document.SaveToFile(output, FileFormat.Docx);

			//Dispose the document
			document.Dispose();

            //Launching the Word file.
            WordDocViewer(output);


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
