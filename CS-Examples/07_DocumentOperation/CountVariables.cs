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

namespace CountVariables
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create Word document.
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_6.docx");

            //Get the number of variables in the document.
            int number = document.Variables.Count;

            StringBuilder content = new StringBuilder();
            content.AppendLine("The number of variables is: " + number.ToString());

            String result = "Result-CountVariables.txt";

            //Save to file.
            File.WriteAllText(result,content.ToString());

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
