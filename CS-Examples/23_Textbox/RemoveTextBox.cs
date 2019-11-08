using System;
using System.Windows.Forms;
using Spire.Doc;

namespace RemoveTextBox
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
            string input = @"..\..\..\..\..\..\Data\TextBoxTemplate.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Remove the first text box
            doc.TextBoxes.RemoveAt(0);

            //Clear all the text boxes
            //Doc.TextBoxes.Clear();

            //Save and launch document
            string output = "RemoveTextBox.docx";
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
