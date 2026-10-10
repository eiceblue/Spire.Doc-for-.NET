Imports System
Imports System.Windows.Forms
Imports Spire.Doc

Namespace CopyMacro
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new Document object
            Dim doc1 As Document = New Document()
            ' Load the source document containing the macro
            doc1.LoadFromFile(@"..\..\..\..\..\..\Data\Macro.docm")

            ' Get the VBA project from the source document
            Dim clonedProject As Spire.Doc.Vba.VbaProject = doc1.VbaProject
            ' Create a new Document object for the target file
            Dim doc2 As Document = New Document()
            ' Load the target document
            doc2.LoadFromFile(@"..\..\..\..\..\..\Data\test.docm")
            ' Assign the VBA project to the target document
            doc2.VbaProject = clonedProject.Clone()
            ' Save the target document as a new Docm file
            doc2.SaveToFile("CloneMacro.docm", Spire.Doc.FileFormat.Docm)
            ' Close the source document
            doc1.Close()
            ' Dispose of the source document to release resources
            doc1.Dispose()
            
             ' Close the source document
            doc2.Close()
            ' Dispose of the source document to release resources
            doc2.Dispose()
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

    End Class
End Namespace
