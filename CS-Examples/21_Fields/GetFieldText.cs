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

namespace GetFieldText
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
			Document document = new Document(@"..\..\..\..\..\..\Data\SampleB_1.docx");

			// Get the collection of fields in the document
			FieldCollection fields = document.Fields;

			// Iterate through each field in the collection
			foreach (Field field in fields)
			{
				// Get the text of the field
				string fieldText = field.FieldText;

				// Append the field text to the StringBuilder
				sb.Append("The field text is \"" + fieldText + "\".\r\n");
			}

			// Write the field information to a text file
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
