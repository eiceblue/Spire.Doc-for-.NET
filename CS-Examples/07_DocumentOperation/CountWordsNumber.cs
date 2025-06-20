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

namespace CountWordsNumber
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

			// Load the document from the specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

			// Create a StringBuilder to store the content
			StringBuilder content = new StringBuilder();

			// Append the character count information to the content
			content.AppendLine("CharCount: " + document.BuiltinDocumentProperties.CharCount);
			content.AppendLine("CharCountWithSpace: " + document.BuiltinDocumentProperties.CharCountWithSpace);

			// Append the word count information to the content
			content.AppendLine("WordCount: " + document.BuiltinDocumentProperties.WordCount);

			// Specify the output file name for the result
			string result = "Result-CountWordsNumber.txt";

			// Write the content to the specified file path
			File.WriteAllText(result, content.ToString());

			// Dispose of the Document object to release resources
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
