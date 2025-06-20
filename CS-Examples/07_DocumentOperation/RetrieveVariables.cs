using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;

namespace RetrieveVariables
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object
			Document document = new Document();

			// Load a Word document from a specific file path
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_6.docx");

			// Get the name of the variable at index 0
			string s1 = document.Variables.GetNameByIndex(0);

			// Get the value of the variable at index 0
			string s2 = document.Variables.GetValueByIndex(0);

			// Get the value of the variable with the name "A1"
			string s3 = document.Variables["A1"];

			// Create a StringBuilder to store the content
			StringBuilder content = new StringBuilder();
			content.AppendLine("The name of the variable retrieved by index 0 is: " + s1);
			content.AppendLine("The value of the variable retrieved by index 0 is: " + s2);
			content.AppendLine("The value of the variable retrieved by name \"A1\" is: " + s3);

			// Specify the output file name
			string result = "Result-RetrieveVariables.txt";

			// Write the content to a text file
			File.WriteAllText(result, content.ToString());

			// Dispose the Document object to release resources
			document.Dispose();           

            //Launch the file.
            WordDocViewer(result);
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
