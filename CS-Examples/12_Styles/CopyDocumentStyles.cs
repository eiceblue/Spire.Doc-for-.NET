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
            //Load source document from disk
            Document srcDoc = new Document();
            srcDoc.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Toc.docx");

            //Load destination document from disk
            Document destDoc= new Document();
            destDoc.LoadFromFile(@"..\..\..\..\..\..\Data\Template_N3.docx");

            //Get the style collections of source document
            Spire.Doc.Collections.StyleCollection styles = srcDoc.Styles;

            //Add the style to destination document
            foreach (Style style in styles)
            {
                destDoc.Styles.Add(style);
            }

            //Save the Word file
            string output = "CopyDocumentStyles_out.docx";
            destDoc.SaveToFile(output, FileFormat.Docx2013);

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
