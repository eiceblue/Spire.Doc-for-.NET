using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Vba;
using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ReadMacro
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object to handle the Word file
            Document doc = new Document();

            // Load the macro-enabled Word document from the specified relative path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Macro.docm");

            // Retrieve the VBA project associated with the document
            VbaProject vbaProject = doc.VbaProject;

            // Get the name of the VBA project
            string name = vbaProject.Name;
            // Get the code page (character encoding) used by the VBA project
            int codePage = vbaProject.CodePage;
            // Check if the VBA project is protected by a password
            bool isProtected = vbaProject.IsProtected;
            // Check if the VBA project is digitally signed
            bool isSigned = vbaProject.IsSigned;

            // Get the total number of VBA modules in the project
            int count = doc.VbaProject.Modules.Count;

            // Initialize a StringBuilder to accumulate the extracted module information
            StringBuilder stringBuilder = new StringBuilder();

            // Loop through each VBA module in the project
            for (int i = 0; i < doc.VbaProject.Modules.Count; i++)
            {
                // Get the current VBA module at the specified index
                VbaModule vbaModule = doc.VbaProject.Modules[i];
                // Retrieve the name of the current module
                string moduleName = vbaModule.Name;
                // Retrieve the source code of the current module
                string sourceCode = vbaModule.SourceCode;
                // Get the type of the current VBA module (e.g., standard, class, document)
                Spire.Doc.Vba.VbaModuleType moduleType = vbaModule.Type;

                // Append the module name to the string builder
                stringBuilder.AppendLine("moduleName: " + moduleName);
                // Append the module's source code to the string builder
                stringBuilder.AppendLine("sourceCode: " + sourceCode);
                // Append the module type to the string builder
                stringBuilder.AppendLine("moduleType: " + moduleType.ToString());
            }

            // Write the accumulated module information to a text file named "ReadMacro.txt"
            File.WriteAllText("ReadMacro.txt", stringBuilder.ToString());
            // Close the document and release the file handle
            doc.Close();
            // Dispose of the document object to free up unmanaged resources
            doc.Dispose();

            WordDocViewer("ReadMacro.txt");
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
