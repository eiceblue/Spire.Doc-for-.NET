using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace DetectAndRemoveVBAMacros
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

            // Load the Word document from the specified file that may contain VBA macros
            document.LoadFromFile(@"..\..\..\..\..\..\Data\DetectAndRemoveVBAMacros.docm");

            // Check if the document contains VBA macros
            if (document.IsContainMacro)
            {
                // Clear/remove the VBA macros from the document
                document.ClearMacros();
            }

            // Specify the name for the resulting document after removing VBA macros
            String result = "Result-DetectAndRemoveVBAMacros.docm";

            // Save the modified document to a new file with the specified name and format (Docm for macro-enabled document)
            document.SaveToFile(result, FileFormat.Docm);

            // Dispose of the document object when finished using it
            document.Dispose();

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
