using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemoveMacro
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
            Document doc = new Document();

            // Load the macro-enabled Word document from the specified file path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Macro.docm");

            // (Optional) Remove a specific VBA module named "SampleModule1"
            //doc.VbaProject.Modules.Remove("SampleModule1");

            // Clear all macros from the document
            doc.ClearMacros();

            // Save the modified document as a new Docm file
            doc.SaveToFile("RemoveMacro.docm", Spire.Doc.FileFormat.Docm);

            // Close the document to release file handles
            doc.Close();

            // Dispose of the document object to free up unmanaged resources
            doc.Dispose();

            //Launching the Word file.
            WordDocViewer("RemoveMacro.docm");
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
