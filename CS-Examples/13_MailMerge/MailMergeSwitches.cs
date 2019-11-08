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
            String input = @"..\..\..\..\..\..\Data\MailMergeSwitches.docx";

            Document doc = new Document();
            //Load a mail merge template file
            doc.LoadFromFile(input);

            string[] fieldName = new string[] { "XX_Name" };
            string[] fieldValue = new string[] { "Jason Tang" };

            doc.MailMerge.Execute(fieldName, fieldValue);
            string result = "MailMergeSwitches_out.docx";
            doc.SaveToFile(result, FileFormat.Docx);
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
