Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Vba
Imports System
Imports System.IO
Imports System.Text
Imports System.Windows.Forms

Namespace ReadMacro
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new Document object to handle the Word file
            Dim doc As Document = New Document()

            ' Load the macro-enabled Word document from the specified relative path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Macro.docm")

            ' Retrieve the VBA project associated with the document
            Dim vbaProject As VbaProject = doc.VbaProject

            ' Get the name of the VBA project
            Dim name As String = vbaProject.Name
            ' Get the code page (character encoding) used by the VBA project
            Dim codePage As Integer = vbaProject.CodePage
            ' Check if the VBA project is protected by a password
            Dim isProtected As Boolean = vbaProject.IsProtected
            ' Check if the VBA project is digitally signed
            Dim isSigned As Boolean = vbaProject.IsSigned

            ' Get the total number of VBA modules in the project
            Dim count As Integer = doc.VbaProject.Modules.Count

            ' Initialize a StringBuilder to accumulate the extracted module information
            Dim stringBuilder As StringBuilder = New StringBuilder()

            ' Loop through each VBA module in the project
            Dim i As Integer = 0
            While i < doc.VbaProject.Modules.Count
                ' Get the current VBA module at the specified index
                Dim vbaModule As VbaModule = doc.VbaProject.Modules[i]
                ' Retrieve the name of the current module
                Dim moduleName As String = vbaModule.Name
                ' Retrieve the source code of the current module
                Dim sourceCode As String = vbaModule.SourceCode
                ' Get the type of the current VBA module (e.g., standard, class, document)
                Dim moduleType As Spire.Doc.Vba.VbaModuleType = vbaModule.Type

                ' Append the module name to the string builder
                stringBuilder.AppendLine("moduleName: " + moduleName)
                ' Append the module's source code to the string builder
                stringBuilder.AppendLine("sourceCode: " + sourceCode)
                ' Append the module type to the string builder
                stringBuilder.AppendLine("moduleType: " + moduleType.ToString())
            Next

            ' Write the accumulated module information to a text file named "ReadMacro.txt"
            File.WriteAllText("ReadMacro.txt", stringBuilder.ToString())
            ' Close the document and release the file handle
            doc.Close()
            ' Dispose of the document object to free up unmanaged resources
            doc.Dispose()

            WordDocViewer("ReadMacro.txt")
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

    End Class
End Namespace
