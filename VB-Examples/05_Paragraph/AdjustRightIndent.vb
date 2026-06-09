Imports System
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AdjustRightIndent
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new instance of the Document class
            Dim doc As Document = New Document()

            ' Add a new section to the document
            Dim section As Section = doc.AddSection()

            ' Add a new paragraph to the body of the section
            Dim paragraph As Paragraph = section.Body.AddParagraph()

            ' Set the text content of the paragraph
            paragraph.Text = "Hello World!"

            ' Enable the adjustment of the right indent for the paragraph format
            paragraph.Format.AdjustRightIndent = True

            ' Add another new paragraph to the body of the section
            paragraph = section.Body.AddParagraph()

            ' Set the text content for the second paragraph
            paragraph.Text = "Thank you for using the Spire.Doc product."

            ' Disable the adjustment of the right indent for this paragraph
            paragraph.Format.AdjustRightIndent = False

            ' Define the file path and name for the output document
            Dim result As String = "AdjustRightIndent.docx"

            ' Save the document to a file in Docx 2016 format
            doc.SaveToFile(result, FileFormat.Docx2016)

            ' Close the document to release resources
            doc.Close()

            ' Dispose of the document object to free up memory
            doc.Dispose()

            'Launching the Word file.
            WordDocViewer(result)
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

    End Class
End Namespace
