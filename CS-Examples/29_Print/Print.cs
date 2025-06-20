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

namespace Print
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
     
            // Create a new instance of Document
            Document document = new Document();

            // Load the Word document from the specified template file
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template.docx");

            // Create a new PrintDialog
            PrintDialog dialog = new PrintDialog();

            // Allow printing of the current page
            dialog.AllowCurrentPage = true;

            // Allow printing of a range of pages
            dialog.AllowSomePages = true;

            // Use the system's default print dialog for selecting printer settings
            dialog.UseEXDialog = true;

            try
            {
                // Set the PrintDialog property of the document to the created PrintDialog
                document.PrintDialog = dialog;

                // Set the PrintDocument property of the PrintDialog to the document's PrintDocument
                dialog.Document = document.PrintDocument;

                // Print the document using the PrintDialog
                dialog.Document.Print();
            }
			
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            // Dispose of the document object when finished using it
            document.Dispose();
        }
    }
}
