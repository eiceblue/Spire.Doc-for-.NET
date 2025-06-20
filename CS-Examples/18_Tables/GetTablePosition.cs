using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;

namespace GetTablePosition
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

            // Load an existing Word document from a file
            document.LoadFromFile(@"..\..\..\..\..\..\Data\TableSample-Az.docx");

            // Get the first section of the document
            Section section = document.Sections[0];

			// Get the first table in the section
			Table table = section.Tables[0] as Table;

			// Create a StringBuilder to store the output content
			StringBuilder stringBuilder = new StringBuilder();

			// Check if text wrapping is enabled around the table
			if (table.Format.WrapTextAround)
			{
				// Get the positioning information for the table
				TablePositioning position = table.Format.Positioning;

				// Append horizontal positioning information to the output content
				stringBuilder.AppendLine("Horizontal:");
				stringBuilder.AppendLine("Position: " + position.HorizPosition + " pt");
				stringBuilder.AppendLine("Absolute Position: " + position.HorizPositionAbs + ", Relative to: " + position.HorizRelationTo);
				stringBuilder.AppendLine();

				// Append vertical positioning information to the output content
				stringBuilder.AppendLine("Vertical:");
				stringBuilder.AppendLine("Position: " + position.VertPosition + " pt");
				stringBuilder.AppendLine("Absolute Position: " + position.VertPositionAbs + ", Relative to: " + position.VertRelationTo);
				stringBuilder.AppendLine();

				// Append distance from surrounding text information to the output content
				stringBuilder.AppendLine("Distance from surrounding text:");
				stringBuilder.AppendLine("Top: " + position.DistanceFromTop + " pt, Left: " + position.DistanceFromLeft + " pt");
				stringBuilder.AppendLine("Bottom: " + position.DistanceFromBottom + " pt, Right: " + position.DistanceFromRight + " pt");
			}

			// Specify the output file path
			string result = "GetTablePosition_out.txt";

			// Write the output content to the output file
			File.WriteAllText(result, stringBuilder.ToString());

			// Dispose of the document object to free up resources
			document.Dispose();
			
            //Launching the Word file.
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
