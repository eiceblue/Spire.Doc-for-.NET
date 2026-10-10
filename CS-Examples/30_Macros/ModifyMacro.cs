using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Vba;
using System;
using System.Windows.Forms;

namespace ModifyMacro
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // Create a new Document object to work with Word files
            Document doc = new Document();

            // Load an existing macro-enabled Word document from the specified relative path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Macro.docm");

            // Retrieve the specific VBA module named "SampleModule2" from the document's VBA project
            VbaModule vbaModule = doc.VbaProject.Modules["SampleModule2"];

            // Replace or set the source code of the VBA module with a new macro that indents a paragraph
            vbaModule.SourceCode = @"
Sub IndentParagraph()
    Selection.ParagraphFormat.LeftIndent = InchesToPoints(2.5)
End Sub";

            // Save the modified document as a new macro-enabled Word file (.docm)
            doc.SaveToFile("ModifyMacro.docm", Spire.Doc.FileFormat.Docm);

            // Close the document to release file handles
            doc.Close();

            // Dispose of the document object to free up unmanaged resources
            doc.Dispose();
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
