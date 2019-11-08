using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace DeleteTableFromTextBox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load the document
            string input = @"..\..\..\..\..\..\Data\TextBoxTable.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the first textbox
            Spire.Doc.Fields.TextBox textbox = doc.TextBoxes[0];

            //Remove the first table from the textbox
            textbox.Body.Tables.RemoveAt(0);

            //Save and launch document
            string output = "DeleteTableFromTextBox.docx";
            doc.SaveToFile(output, FileFormat.Docx);
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
