using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CellMergeStatus
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\CellMergeStatus.docx";

			//Create a Word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the first section
			Section section = doc.Sections[0];

			//Get the first table in the section
			Table table = section.Tables[0] as Table;

			//Create a StringBuilder instance
			StringBuilder stringBuidler = new StringBuilder();

			//Loop through the table rows
			for (int i = 0; i < table.Rows.Count; i++)
			{
				//Get the table rows
				TableRow tableRow = table.Rows[i];

				//Loop through the cells of the row
				for (int j = 0; j < tableRow.Cells.Count; j++)
				{
					//Get each cell
					TableCell tableCell = tableRow.Cells[j];

					//Returns the way of vertical merging of the cell
					CellMerge verticalMerge = tableCell.CellFormat.VerticalMerge;

					//Get the status of cell merge 
					short horizontalMerge = tableCell.GridSpan;
					if (verticalMerge == CellMerge.None && horizontalMerge == 1)
					{
						stringBuidler.Append("Row " + i + ", cell " + j + ": ");
						stringBuidler.AppendLine("This cell isn't merged.");
					}
					else
					{
						stringBuidler.Append("Row " + i + ", cell " + j + ": ");
						stringBuidler.AppendLine("This cell is merged.");
					}
				}

				//Append an empty line
				stringBuidler.AppendLine();
			}

			//Save to a text file
			string output = "CellMergeStatus.txt";
			File.WriteAllText(output, stringBuidler.ToString());

			//Dispose the document
			doc.Dispose();
			
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
