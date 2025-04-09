using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Pages;

namespace FixedLayout
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Specify the input file path
            string inputFile = @"..\..\..\..\..\..\Data\Template_Docx_3.docx";

            // Create a new instance of Document
            Document doc = new Document();

            // Load the document from the specified file
            doc.LoadFromFile(inputFile, FileFormat.Docx);

            // Create a FixedLayoutDocument from the loaded document
            FixedLayoutDocument layoutDoc = new FixedLayoutDocument(doc);

            // Get the first line in the first column of the first page
            FixedLayoutLine line = layoutDoc.Pages[0].Columns[0].Lines[0];

            // Create a StringBuilder to store the output text
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("Line: " + line.Text);

            // Get the paragraph that contains the line and append its text to the StringBuilder
            Paragraph para = line.Paragraph;
            stringBuilder.AppendLine("Paragraph text: " + para.Text);

            // Get the text content of the first page
            string pageText = layoutDoc.Pages[0].Text;
            stringBuilder.AppendLine(pageText);

            // Iterate through each page in the FixedLayoutDocument
            foreach (FixedLayoutPage page in layoutDoc.Pages)
            {
                // Get all the lines on the current page
                LayoutCollection<LayoutElement> lines = page.GetChildEntities(LayoutElementType.Line, true);

                // Append the page index and number of lines to the StringBuilder
                stringBuilder.AppendLine("Page " + page.PageIndex + " has " + lines.Count + " lines.");
            }

            // Append the lines of the first paragraph to the StringBuilder
			// (except runs and nodes in the header and footer).
            stringBuilder.AppendLine("The lines of the first paragraph:");
            foreach (FixedLayoutLine paragraphLine in layoutDoc.GetLayoutEntitiesOfNode(((Section)doc.FirstChild).Body.Paragraphs[0]))
            {
                stringBuilder.AppendLine(paragraphLine.Text.Trim());
                stringBuilder.AppendLine(paragraphLine.Rectangle.ToString());
            }

            // Write the contents of the StringBuilder to a text file
            File.WriteAllText("page.txt", stringBuilder.ToString());

            // Dispose of the document object when finished using it
            doc.Dispose();
        }
    }
}
