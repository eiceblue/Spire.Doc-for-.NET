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
			// Create a StringBuilder to hold the field information
			StringBuilder sb = new StringBuilder();

			// Load the document from a file
			Document document = new Document(@"..\..\..\..\..\..\Data\MailMerge.doc");

			// Get the array of merge field names in the document
			string[] fieldNames = document.MailMerge.GetMergeFieldNames();

			// Append the count of merge fields in the document to the StringBuilder
			sb.Append("The document has " + fieldNames.Length.ToString() + " merge fields.");

			// Append a header for the merge field names
			sb.Append(" The below is the name of the merge field:" + "\r\n");

			// Iterate through each merge field name and append it to the StringBuilder
			foreach (string name in fieldNames)
			{
				sb.AppendLine(name);
			}

			// Write the result to a text file
			File.WriteAllText("result.txt", sb.ToString());

			// Dispose the document object
			document.Dispose();

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
