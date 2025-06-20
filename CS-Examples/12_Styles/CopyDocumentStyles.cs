using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CopyDocumentStyles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a Word document.
			Document srcDoc = new Document();

			//Load the file from disk.
			srcDoc.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Toc.docx");

			//Create another Word document
			Document destDoc = new Document();

			//Load destination document from disk
			destDoc.LoadFromFile(@"..\..\..\..\..\..\Data\Template_N3.docx");

			//Get the style collections of source document
			Spire.Doc.Collections.StyleCollection styles = srcDoc.Styles;

			//Loop throughthe styles of source document
			foreach (Style style in styles)
			{
				//Add the style to destination document
				destDoc.Styles.Add(style);
			}

			//Save the Word file
			string output = "CopyDocumentStyles_out.docx";
			destDoc.SaveToFile(output, FileFormat.Docx2013);

			//Dispose the document
			srcDoc.Dispose();
			destDoc.Dispose();

            //Launch the file
            FileViewer(output);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}
