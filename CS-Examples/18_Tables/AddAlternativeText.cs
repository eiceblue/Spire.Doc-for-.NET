﻿using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddAlternativeText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			string input = @"..\..\..\..\..\..\Data\TableSample.docx";

			//Create a Word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the first section
			Section section = doc.Sections[0];

			//Get the first table in the section
			Table table = section.Tables[0] as Table;

			//Set the table title
			table.Title = "Table 1";
			
			//Add description
			table.TableDescription = "Description Text";

			//Save the document
			string output = "AddAlternativeText.docx";
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
