Imports System
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports Spire.Doc.Fields

Namespace SetParagraphTextDirection
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Initialize a new Document object.
            Dim doc As Document = New Document()

            ' Add a new section to the document.
            Dim section As Section = doc.AddSection()

            ' Add a new paragraph to the section.
            Dim paragraph As Paragraph = section.AddParagraph()

            ' Append the text "Welcome to China." to the paragraph and get the TextRange object.
            Dim farEastLayout As TextRange = paragraph.AppendText("Welcome to China.")

            ' Create a new FarEastLayout object to define vertical text settings.
            Dim style As FarEastLayout = New FarEastLayout()

            ' Enable vertical text orientation for the layout style.
            style.Vertical = True

            ' Apply the vertical FarEastLayout style to the character format of the text range.
            farEastLayout.CharacterFormat.FarEastLayout = style

            ' Define the output file name for the saved document.
            Dim outputFile As String = "SetParagraphTextDirection.docx"

            ' Save the document to the specified file in DOCX format.
            doc.SaveToFile(outputFile, FileFormat.Docx)

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
