Imports System
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AddSingleLevelList
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new instance of the Document class to represent a Word document
            Dim document As Document = New Document()

            ' Add a new section to the document, which acts as a container for content like paragraphs and tables
            Dim section As Section = document.AddSection()

            ' Define a list template that uses Arabic numerals (1, 2, 3) followed by a dot
            Dim template As ListTemplate = ListTemplate.NumberArabicDot

            ' Register this single-level numbered list template with the document and get a reference to it
            Dim listRef As ListDefinitionReference = document.ListReferences.AddSingleLevelList(template)

            ' Create a new paragraph object within the current section
            Dim paragraph As Paragraph = section.AddParagraph()

            ' Append the text to the newly created paragraph
            paragraph.AppendText("List Item 1")

            ' Apply the previously defined numbered list format (listRef) at level 0 to this paragraph
            paragraph.ListFormat.ApplyListRef(listRef, 0)

            ' Reassign the paragraph variable by adding another new paragraph to the section
            paragraph = section.AddParagraph()

            ' Append the text to this new paragraph
            paragraph.AppendText("List Item 2")

            ' Apply the same numbered list format at level 0 to continue the sequence
            paragraph.ListFormat.ApplyListRef(listRef, 0)

            ' Create a third paragraph in the section for the next list item
            paragraph = section.AddParagraph()

            ' Append the text to the paragraph
            paragraph.AppendText("List Item 3")

            ' Apply the numbered list format at level 0 to complete the list
            paragraph.ListFormat.ApplyListRef(listRef, 0)

            Dim result As String = "addSingleLevelList.docx"
            ' Save the document to a file using Docx format
            document.SaveToFile(result, FileFormat.Docx)

            ' Close the document to release system resources associated with the file
            document.Close()

            ' Dispose of the document object to free up memory
            document.Dispose()

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
