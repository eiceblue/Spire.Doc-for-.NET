Imports System
Imports System.Windows.Forms
Imports Spire.Doc

Namespace MarkdownToDocxUsingTemplateStyles
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Initialize a new Document object by loading the Markdown file from the specified relative path.
            Dim doc As Document = New Document(@"..\..\..\..\..\..\Data\sample.md")

            ' Copy all styles from the specified Word template into the current document.
            doc.CopyStylesFromTemplate(@"..\..\..\..\..\..\Data\template.docx")

            ' Define the output filename for the converted Word document.
            Dim outputFile As String = "MarkdownToDocxUsingTemplateStyles.docx"

            ' Save the processed document to the specified file in DOCX 2016 format.
            doc.SaveToFile(outputFile, FileFormat.Docx2016)

            ' Close the document to release resources.
            doc.Close()

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
