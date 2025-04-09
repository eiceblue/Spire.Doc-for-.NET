Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text

Namespace ReplaceTextWithMergeField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a file
			Dim document As New Document("..\..\..\..\..\..\Data\SampleB_2.docx")

			' Find the text "Test" in the document
			Dim ts As TextSelection = document.FindString("Test", True, True)

			' Get the selected text as a single range
			Dim tr As TextRange = ts.GetAsOneRange()

			' Get the paragraph that contains the selected text
			Dim par As Paragraph = tr.OwnerParagraph

			' Get the index of the selected text within its parent paragraph
			Dim index As Integer = par.ChildObjects.IndexOf(tr)

			' Create a new merge field
			Dim field As New MergeField(document)
			field.FieldName = "MergeField"

			' Insert the merge field at the same position as the selected text
			par.ChildObjects.Insert(index, field)

			' Remove the selected text from the paragraph
			par.ChildObjects.Remove(tr)

			' Save the modified document to a new file
			document.SaveToFile("result.docx", FileFormat.Docx)

			' Dispose of the document object
			document.Dispose()
			
			'Launch result file
			WordDocViewer("result.docx")

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
