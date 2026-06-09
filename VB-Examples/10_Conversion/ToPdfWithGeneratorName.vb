Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Doc

Namespace ToPdfWithGeneratorName
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new Document instance
            Dim document As Document = New Document()

            ' Load the Word document from the specified file path
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\ConvertedTemplate.docx")

            ' Create a ToPdfParameterList instance to configure PDF conversion options
            Dim toPdf As ToPdfParameterList = New ToPdfParameterList()

            ' Define the generator name
            toPdf.GeneratorName = "Spire.Doc for .NET Product"
            document.SaveToFile("ToPdfWithGeneratorName.pdf", toPdf)
            document.Close()
            document.Dispose()

            'view the PDF file.
            WordDocViewer("ToPdfWithGeneratorName.pdf")
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub
    End Class
End Namespace
