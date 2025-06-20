using System;
using System.Windows.Forms;
using Spire.Doc;

namespace ToPdfPreserveFormFields
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			// Load the document from a file
			Document document = new Document(@"..\..\..\..\..\..\..\Data\ToPdfPreserveFormFields.docx");

            // Preserve form field when converting to Pdf
            ToPdfParameterList ppl = new ToPdfParameterList();
            ppl.PreserveFormFields = true;

            document.SaveToFile("ToPdfPreserveFormFields_output.pdf",ppl);
			// Dispose the document object
			document.Dispose();

            //Launch result file
            WordDocViewer("ToPdfPreserveFormFields_output.pdf");

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
