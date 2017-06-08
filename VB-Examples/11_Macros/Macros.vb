Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace Macros
    Partial Public Class Form1
        Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            Dim document As New Document()

            'Loading documetn with macros.
            document.LoadFromFile("../../../../../../Data/Macros.docm", FileFormat.Docm)

            'Removes the macros from the document.
            document.ClearMacros()

            'Save docm file.
            document.SaveToFile("Sample.docm", FileFormat.Docm)

            'Launching the MS Word file.
            WordDocViewer("Sample.docm")
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                Process.Start(fileName)
            Catch
            End Try
        End Sub

    End Class
End Namespace
