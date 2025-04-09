Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace CreateCrossReference
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Add a section to the document
			Dim section As Section = document.AddSection()

			' Add a paragraph to the section and append a bookmark with the specified name
			Dim paragraph As Paragraph = section.AddParagraph()
			paragraph.AppendBookmarkStart("MyBookmark")
			paragraph.AppendText("Text inside a bookmark")
			paragraph.AppendBookmarkEnd("MyBookmark")

			' Add line breaks to the paragraph
			For i As Integer = 0 To 3
				paragraph.AppendBreak(BreakType.LineBreak)
			Next i

			' Create a new Field object for referencing the bookmark
			Dim field As New Field(document)
			field.Type = FieldType.FieldRef
			field.Code = "REF MyBookmark \p \h"

			' Add a new paragraph to the section and append text and the field
			paragraph = section.AddParagraph()
			paragraph.AppendText("For more information, see ")
			paragraph.ChildObjects.Add(field)

			' Add a field separator to the paragraph
			Dim fieldSeparator As New FieldMark(document, FieldMarkType.FieldSeparator)
			paragraph.ChildObjects.Add(fieldSeparator)

			' Create a TextRange object and set its text
			Dim tr As New TextRange(document)
			tr.Text = "above"
			paragraph.ChildObjects.Add(tr)

			' Add a field end mark to the paragraph
			Dim fieldEnd As New FieldMark(document, FieldMarkType.FieldEnd)
			paragraph.ChildObjects.Add(fieldEnd)

			' Specify the file name for the result document
			Dim result As String = "Result-CreateCrossReferenceToBookmark.docx"

			' Save the document to a file
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose the document object
			document.Dispose()

			'Launch the MS Word file.
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
