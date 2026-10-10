Imports System
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemoveMacro
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new Document instance
            Dim doc As Document = New Document()

            ' Load the macro-enabled Word document from the specified file path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Macro.docm")

            ' (Optional) Remove a specific VBA module named "SampleModule1"
            'doc.VbaProject.Modules.Remove("SampleModule1");

            ' Clear all macros from the document
            doc.ClearMacros()

            ' Save the modified document as a new Docm file
            doc.SaveToFile("RemoveMacro.docm", Spire.Doc.FileFormat.Docm)

            ' Close the document to release file handles
            doc.Close()

            ' Dispose of the document object to free up unmanaged resources
            doc.Dispose()

            'Launching the Word file.
            WordDocViewer("RemoveMacro.docm")
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

    End Class
End Namespace
