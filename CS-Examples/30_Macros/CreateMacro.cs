using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Vba;

namespace CreateMacro
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Word document instance
            Document document = new Document();

            // Add a new section to the document
            Section section = document.AddSection();

            // Add a new paragraph to the section
            Paragraph paragraph = section.AddParagraph();

            // Append text content to the paragraph
            paragraph.AppendText("There is a VBA Macro. ");

            // Create a new VBA project to hold macros
            VbaProject vbaProject = new VbaProject();

            // Set the name of the VBA project
            vbaProject.Name = "SampleVBAMacro";

            // Assign the VBA project to the document
            document.VbaProject = vbaProject;

            // Add a standard VBA module named "SampleModule1" to the project
            VbaModule vbaModule1 = document.VbaProject.Modules.Add("SampleModule1", Spire.Doc.Vba.VbaModuleType.StdModule);

            // Define the VBA source code for the first module (displays document info and writes "Hello World!")
            vbaModule1.SourceCode = @"
 Sub DocumnetInfo()
     MsgBox ""create time: "" &Now()
     MsgBox ""Pages:"" & ActiveDocument.Range.ComputeStatistics(wdStatisticPages)
 End Sub

 Sub WriteHello()
     Selection.TypeText Text:=""Hello World!""
 End Sub";

            // Add a second standard VBA module named "SampleModule2" to the project
            VbaModule vbaModule2 = document.VbaProject.Modules.Add("SampleModule2", Spire.Doc.Vba.VbaModuleType.StdModule);

            // Define the VBA source code for the second module (inserts current date and indents a paragraph)
            vbaModule2.SourceCode = @"
 Sub InsertCurrentDate()
     Selection.TypeText Text:=Format(Now(),""yyyy-mm-dd hh:mm:ss"")
 End Sub

 Sub IndentParagraph()
     Selection.ParagraphFormat.LeftIndent = InchesToPoints(0.5)
 End Sub";

            // Save the document as a macro-enabled Word file (.docm)
            document.SaveToFile("CreateMacro.docm", Spire.Doc.FileFormat.Docm);

            // Close the document and release resources
            document.Close();

            // Dispose of the document object and release any associated resources
            document.Dispose();

            //Launching the Word file.
            WordDocViewer("CreateMacro.docm");
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
