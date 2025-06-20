using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;
using Spire.Doc.Fields;
using System.Data;
using System.Data.OleDb;
namespace MailMergeSwitches
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\MailMergeSwitches.docx";

			//Create a Word document
			Document doc = new Document();

			//Load a mail merge template file
			doc.LoadFromFile(input);

			//Define the field names for the mail merge
			string[] fieldName = new string[] { "XX_Name" };

			//Define the field values for the mail merge
			string[] fieldValue = new string[] { "Jason Tang" };

			//Execute the mail merge using the field names and values
			doc.MailMerge.Execute(fieldName, fieldValue);

			//Save to file
			string result = "MailMergeSwitches_out.docx";
			doc.SaveToFile(result, FileFormat.Docx);

			// Dispose the document object
			doc.Dispose();
            WordViewer(result);
        }      
        private void WordViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
