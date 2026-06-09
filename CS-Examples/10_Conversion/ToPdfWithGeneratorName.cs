using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;

namespace ToPdfWithGeneratorName
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document instance
            Document document = new Document();

            // Load the Word document from the specified file path
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\ConvertedTemplate.docx");

            // Create a ToPdfParameterList instance to configure PDF conversion options
            ToPdfParameterList toPdf = new ToPdfParameterList();

            // Define the generator name
            toPdf.GeneratorName = "Spire.Doc for .NET Product";
            document.SaveToFile("ToPdfWithGeneratorName.pdf", toPdf);
            document.Close();
            document.Dispose();

            //view the PDF file.
            WordDocViewer("ToPdfWithGeneratorName.pdf");
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
