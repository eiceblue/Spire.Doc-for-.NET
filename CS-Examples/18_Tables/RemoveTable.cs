using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemoveTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string input = @"..\..\..\..\..\..\Data\Template.docx";
            // Create a new Document object
		    Document doc = new Document();

		    // Load an existing Word document from a file
            doc.LoadFromFile(input);

            //Remove the first Table            
            doc.Sections[0].Tables.RemoveAt(0);

            //Save the document
            string output = "RemoveTable.docx";
            doc.SaveToFile(output, FileFormat.Docx);
			
			//Dispose of the document object to free up resources
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
