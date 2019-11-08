using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Collections;
using System.Text;

namespace GetMergeFieldName
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            //Open a Word document
            Document document = new Document(@"..\..\..\..\..\..\Data\MailMerge.doc");

            //Get merge field name
            string[] fieldNames = document.MailMerge.GetMergeFieldNames();

            sb.Append("The document has " + fieldNames.Length.ToString() + " merge fields.");
            sb.Append(" The below is the name of the merge field:"+"\r\n");
            foreach (string name in fieldNames)
            {
                sb.AppendLine(name);
            }

            File.WriteAllText("result.txt", sb.ToString());

            //Launch result file
            WordDocViewer("result.txt");

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
