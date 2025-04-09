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

			' Remove the text range from its owning paragraph.
			ownParagraph.ChildObjects.RemoveAt(rangeIndex)

			' Create a temporary list to store the removed ChildObjects.
			Dim tempList As New List(Of DocumentObject)()

			' Clone and remove ChildObjects starting from the range index until the end of the paragraph.
			Dim i As Integer = rangeIndex
			Do While i < ownParagraph.ChildObjects.Count
				tempList.Add(ownParagraph.ChildObjects(rangeIndex).Clone())
				ownParagraph.ChildObjects.RemoveAt(rangeIndex)
				i += 1
			Loop

			' Append a field named "MyFieldName" of type FieldType.FieldMergeField to the owning paragraph.
			ownParagraph.AppendField("MyFieldName", FieldType.FieldMergeField)

			' Add back the cloned ChildObjects to the owning paragraph.
			For Each obj As DocumentObject In tempList
				ownParagraph.ChildObjects.Add(obj)
			Next obj

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
