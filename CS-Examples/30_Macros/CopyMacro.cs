using System;
using System.Windows.Forms;
using Spire.Doc;

namespace CopyMacro
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object
            Document doc1 = new Document();
            // Load the source document containing the macro
            doc1.LoadFromFile(@"..\..\..\..\..\..\Data\Macro.docm");

            // Get the VBA project from the source document
            Spire.Doc.Vba.VbaProject clonedProject = doc1.VbaProject;
            // Create a new Document object for the target file
            Document doc2 = new Document();
            // Load the target document
            doc2.LoadFromFile(@"..\..\..\..\..\..\Data\test.docm");
            // Assign the VBA project to the target document
            doc2.VbaProject = clonedProject.Clone();
            // Save the target document as a new Docm file
            doc2.SaveToFile("CloneMacro.docm", Spire.Doc.FileFormat.Docm);
            // Close the source document
            doc1.Close();
            // Dispose of the source document to release resources
            doc1.Dispose();
            
             // Close the source document
            doc2.Close();
            // Dispose of the source document to release resources
            doc2.Dispose();
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
