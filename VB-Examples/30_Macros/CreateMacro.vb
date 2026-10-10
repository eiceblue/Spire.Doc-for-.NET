Imports System
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Vba

Namespace CreateMacro
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new Word document instance
            Dim document As Document = New Document()

            ' Add a new section to the document
            Dim section As Section = document.AddSection()

            ' Add a new paragraph to the section
            Dim paragraph As Paragraph = section.AddParagraph()

            ' Append text content to the paragraph
            paragraph.AppendText("There Is a VBA Macro. ")

            ' Create a new VBA project to hold macros
            Dim vbaProject As VbaProject = New VbaProject()

            ' Set the name of the VBA project
            vbaProject.Name = "SampleVBAMacro"

            ' Assign the VBA project to the document
            document.VbaProject = vbaProject

            ' Add a standard VBA module named "SampleModule1" to the project
            Dim vbaModule1 As VbaModule = document.VbaProject.Modules.Add("SampleModule1", Spire.Doc.Vba.VbaModuleType.StdModule)

            ' Define the VBA source code for the first module (displays document info and writes "Hello World!")
            vbaModule1.SourceCode = @"
 Sub DocumnetInfo()
     MsgBox ""create time: "" &Now()
     MsgBox ""Pages:"" & ActiveDocument.Range.ComputeStatistics(wdStatisticPages)
 End Sub

 Sub WriteHello()
     Selection.TypeText Text:=""Hello World!""
 End Sub"

            ' Add a second standard VBA module named "SampleModule2" to the project
            Dim vbaModule2 As VbaModule = document.VbaProject.Modules.Add("SampleModule2", Spire.Doc.Vba.VbaModuleType.StdModule)

            ' Define the VBA source code for the second module (inserts current date and indents a paragraph)
            vbaModule2.SourceCode = @"
 Sub InsertCurrentDate()
     Selection.TypeText Text:=Format(Now(),""yyyy-mm-dd hh:mm:ss"")
 End Sub

 Sub IndentParagraph()
     Selection.ParagraphFormat.LeftIndent = InchesToPoints(0.5)
 End Sub"

            ' Save the document as a macro-enabled Word file (.docm)
            document.SaveToFile("CreateMacro.docm", Spire.Doc.FileFormat.Docm)

            ' Close the document and release resources
            document.Close()

            ' Dispose of the document object and release any associated resources
            document.Dispose()

            'Launching the Word file.
            WordDocViewer("CreateMacro.docm")
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

    End Class
End Namespace
