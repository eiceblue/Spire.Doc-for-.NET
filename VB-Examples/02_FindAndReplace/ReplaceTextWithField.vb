Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ReplaceTextWithField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object.
			Dim document As New Document()

			' Load the Word document from the specified file path.
			document.LoadFromFile("..\..\..\..\..\..\Data\ReplaceTextWithField.docx")

			' Find the first occurrence of the word "summary" in the document and retrieve it as a TextSelection object.
			Dim selection As TextSelection = document.FindString("summary", False, True)

			' Convert the selected text into a single TextRange.
			Dim textRange As TextRange = selection.GetAsOneRange()

			' Get the paragraph that owns the text range.
			Dim ownParagraph As Paragraph = textRange.OwnerParagraph

			' Retrieve the index of the text range within its owning paragraph.
			Dim rangeIndex As Integer = ownParagraph.ChildObjects.IndexOf(textRange)

            'Add the date field
            Dim eqField As New Field(document)
            eqField.Type = FieldType.FieldDate
            eqField.Code = "DATE  \@" & """yyyyMMdd """
            ownParagraph.ChildObjects.Insert(rangeIndex, eqField)
            Dim mark As New FieldMark(document, FieldMarkType.FieldSeparator)
            ownParagraph.ChildObjects.Insert(rangeIndex + 1, mark)
            Dim [end] As New FieldMark(document, FieldMarkType.FieldEnd)
            ownParagraph.ChildObjects.Insert(rangeIndex + 2, [end])

            ' Remove the selected text from the paragraph
            ownParagraph.ChildObjects.Remove(textRange)

			' Specify the output file name for saving the modified document.
			Dim output As String = "ReplaceTextWithField_output.docx"

			' Save the modified document to the specified output file in Docx format.
			document.SaveToFile(output, FileFormat.Docx)

			' Dispose of the Document object to release resources.
			document.Dispose()
			
			WordDocViewer(output)
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
