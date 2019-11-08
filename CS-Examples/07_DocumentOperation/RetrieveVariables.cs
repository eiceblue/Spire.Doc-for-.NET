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

namespace RetrieveVariables
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

            //Retrieve name of the variable by index.
            string s1 = document.Variables.GetNameByIndex(0);

            //Retrieve value of the variable by index.
            string s2 = document.Variables.GetValueByIndex(0);

            //Retrieve the value of the variable by name.
            string s3 = document.Variables["A1"];

            StringBuilder content = new StringBuilder();
            content.AppendLine("The name of the variable retrieved by index 0 is: " + s1);
            content.AppendLine("The vaule of the variable retrieved by index 0 is: " + s2);
            content.AppendLine("The vaule of the variable retrieved by name \"A1\" is: " + s3);

            String result = "Result-RetrieveVariables.txt";

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
