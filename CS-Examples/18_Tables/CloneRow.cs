using System;
using System.Windows.Forms;
using Spire.Doc;

namespace CloneRow
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			string input = @"..\..\..\..\..\..\Data\TableTemplate.docx";

			//Create a Word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the first section
			Section se = doc.Sections[0];

			//Get the first row of the first table
			TableRow firstRow = se.Tables[0].Rows[0];

			//Copy the first row to clone_FirstRow via TableRow.clone()
			TableRow clone_FirstRow = firstRow.Clone();

			//Add a table row to collection
			se.Tables[0].Rows.Add(clone_FirstRow);

			//Save and launch document
			string output = "CloneRow_output.docx";
			doc.SaveToFile(output, FileFormat.Docx);

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
