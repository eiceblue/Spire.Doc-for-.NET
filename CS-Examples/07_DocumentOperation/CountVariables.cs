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

namespace CountVariables
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

			// Load a Word document from the specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_6.docx");

			// Get the number of variables present in the document
			int number = document.Variables.Count;

			// Create a StringBuilder to hold the content for the result file
			StringBuilder content = new StringBuilder();

			// Append the number of variables to the content
			content.AppendLine("The number of variables is: " + number.ToString());

			// Specify the file name for the saved result file
			string result = "Result-CountVariables.txt";

			// Write the content to a text file
			File.WriteAllText(result, content.ToString());

			// Release the resources used by the document
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
