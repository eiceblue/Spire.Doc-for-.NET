using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Vba;

namespace AddVBAMacro
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize a new Document object
            Document doc = new Document();

            // Add a new section, then a paragraph to that section, and append the text "Add VBA macro"
            doc.AddSection().AddParagraph().AppendText("Add VBA macro");

            // Create a new VBA project instance
            VbaProject vbaProject = new VbaProject();

            // Set the name of the VBA project
            vbaProject.Name = "SampleVBAMacro";

            // Assign the created VBA project to the document
            doc.VbaProject = vbaProject;

            // Add a new standard VBA module named "SampleModule1" to the project
            VbaModule vbaModule1 = doc.VbaProject.Modules.Add("SampleModule1", VbaModuleType.StdModule);

            // Define the source code for the first module containing two macros:
            vbaModule1.SourceCode = @"
                Sub DocumnetInfo()
                    MsgBox ""create time: "" &Now()
                    MsgBox ""Pages:"" & ActiveDocument.Range.ComputeStatistics(wdStatisticPages)
                End Sub

                Sub WriteHello()
                    Selection.TypeText Text:=""Hello World!""
                End Sub";

            // Add a second standard VBA module named "SampleModule2" to the project
            VbaModule vbaModule2 = doc.VbaProject.Modules.Add("SampleModule2", VbaModuleType.StdModule);

            // Define the source code for the second module containing two macros:
            vbaModule2.SourceCode = @"
                Sub InsertCurrentDate()
                    Selection.TypeText Text:=Format(Now(),""yyyy-mm-dd hh:mm:ss"")
                End Sub

                Sub IndentParagraph()
                    Selection.ParagraphFormat.LeftIndent = InchesToPoints(0.5)
                End Sub";

            // Define the output file name with the .docm extension (required for documents containing macros)
            String outputFile = "AddVBAMacro.docm";

            // Save the document as a Macro-Enabled DOCX file (DOCX 2019 format with macros)
            doc.SaveToFile(outputFile, FileFormat.Docm);

            // Close the document to release file handles
            doc.Close();

            // Dispose of the document object to free up memory
            doc.Dispose();

            WordDocViewer(outputFile);
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
