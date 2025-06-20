using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace GetVariables
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

			// Create a StringBuilder to hold the content for the output file
			StringBuilder stringBuilder = new StringBuilder();

			// Append an introductory line to the StringBuilder
			stringBuilder.AppendLine("This document has following variables:");

			// Iterate through each key-value pair in the document's Variables collection
			foreach (KeyValuePair<string, string> entry in document.Variables)
			{
				// Extract the name and value from the current entry
				string name = entry.Key;
				string value = entry.Value;

			// Append the name and value to the StringBuilder
			stringBuilder.AppendLine("Name: " + name + ", " + "Value: " + value);
			}

			// Specify the file name for the saved output file
			string result = "GetVariables_out.txt";

			// Write the content of the StringBuilder to a text file
			File.WriteAllText(result, stringBuilder.ToString());

			// Release the resources used by the document
			document.Dispose();

            WordDocViewer("GetVariables_out.txt");
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
