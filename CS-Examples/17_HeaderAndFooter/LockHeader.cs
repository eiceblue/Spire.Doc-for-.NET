﻿using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace LockHeader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\HeaderAndFooter.docx";

			//Create a word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the first section
			Section section = doc.Sections[0];

			//Protect the document and set the ProtectionType as AllowOnlyFormFields
			doc.Protect(ProtectionType.AllowOnlyFormFields, "123");

			//Set the ProtectForm as false to unprotect the section
			section.ProtectForm = false;

			//Save the document
			string output = "LockHeader.docx";
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
