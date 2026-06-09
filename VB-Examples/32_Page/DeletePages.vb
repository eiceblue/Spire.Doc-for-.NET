Imports System
Imports System.Windows.Forms
Imports Spire.Doc

Namespace DeletePages
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Initialize a new Document object
            Dim document As Document = New Document()

            ' Load an existing Word document from the specified relative file path
            document.LoadFromFile(@"..\..\..\..\..\..\Data\RemovePages.docx")

            ' Remove all blank pages from the document
            document.RemoveBlankPages()

            ' Remove specific pages by index (0-based). Here, it removes the 3rd page (index 2) and the 5th page (index 4).
            document.RemovePages(New System.Collections.Generic.List<Integer> { 2, 4 })

            ' Define the output file name for the modified document
            Dim outputFile As String = "DeletePages.docx"

            ' Save the document to the specified file in DOCX 2019 format
            document.SaveToFile(outputFile, FileFormat.Docx2019)

            ' Close the document to release file handles
            document.Close()

            ' Dispose of the document object to free up memory
            document.Dispose()

            'Launch the Word file.
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
