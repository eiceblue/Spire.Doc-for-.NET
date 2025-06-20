using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using System.IO;
using Spire.Doc.Documents;
using Spire.Doc.Pages;

namespace GetPageNumberOfText
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
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

			// Find all occurrences of the string "Spire" in the document
			TextSelection[] textSelections = document.FindAllString("Spire", false, false);

			// Create a FixedLayoutDocument object using the loaded document
			FixedLayoutDocument layoutDoc = new FixedLayoutDocument(document);

			// Initialize a counter for matched words
			int count = 1;

			// Create a StringBuilder to store the result
			StringBuilder builder = new StringBuilder();

			// Iterate through each TextSelection
			foreach (TextSelection selection in textSelections)
			{
				// Get the layout entities for the current selection
				foreach (FixedLayoutSpan line in layoutDoc.GetLayoutEntitiesOfNode(selection.GetRanges()[0]))
				{
					// Get the page index where the matched word is located
					int index = line.PageIndex;

					// Append the result to the StringBuilder
					builder.AppendLine("The matched word " + count + " is on page:" + index);

					// Increment the counter
					count++;
				}
			}

			// Write the result to a text file named "result.txt"
			File.WriteAllText("result.txt", builder.ToString());

			// Dispose the Document object to release resources
			document.Dispose();
			
            System.Diagnostics.Process.Start("result.txt");
        }

    }
}
