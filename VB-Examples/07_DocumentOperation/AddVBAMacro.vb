Imports System
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Vba

Namespace AddVBAMacro
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Initialize a new Document object
            Dim doc As Document = New Document()

            ' Add a new section, then a paragraph to that section, and append the text "Add VBA macro"
            doc.AddSection().AddParagraph().AppendText("Add VBA macro")

            ' Create a new VBA project instance
            Dim vbaProject As VbaProject = New VbaProject()

            ' Set the name of the VBA project
            vbaProject.Name = "SampleVBAMacro"

            ' Assign the created VBA project to the document
            doc.VbaProject = vbaProject

            ' Add a new standard VBA module named "SampleModule1" to the project
            Dim vbaModule1 As VbaModule = doc.VbaProject.Modules.Add("SampleModule1", VbaModuleType.StdModule)

            ' Define the source code for the first module containing two macros:
            vbaModule1.SourceCode = @"
                Sub DocumnetInfo()
                    MsgBox ""create time: "" &Now()
                    MsgBox ""Pages:"" & ActiveDocument.Range.ComputeStatistics(wdStatisticPages)
                End Sub

                Sub WriteHello()
                    Selection.TypeText Text:=""Hello World!""
                End Sub"

            ' Add a second standard VBA module named "SampleModule2" to the project
            Dim vbaModule2 As VbaModule = doc.VbaProject.Modules.Add("SampleModule2", VbaModuleType.StdModule)

            ' Define the source code for the second module containing two macros:
            vbaModule2.SourceCode = @"
                Sub InsertCurrentDate()
                    Selection.TypeText Text:=Format(Now(),""yyyy-mm-dd hh:mm:ss"")
                End Sub

                Sub IndentParagraph()
                    Selection.ParagraphFormat.LeftIndent = InchesToPoints(0.5)
                End Sub"

            ' Define the output file name with the .docm extension (required for documents containing macros)
            Dim outputFile As String = "AddVBAMacro.docm"

            ' Save the document as a Macro-Enabled DOCX file (DOCX 2019 format with macros)
            doc.SaveToFile(outputFile, FileFormat.Docm)

            ' Close the document to release file handles
            doc.Close()

            ' Dispose of the document object to free up memory
            doc.Dispose()

            WordDocViewer(outputFile)
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

    End Class
End Namespace
