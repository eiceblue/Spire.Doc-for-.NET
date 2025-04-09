Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting.Revisions
Imports Spire.Doc.Fields
Imports System.IO

Namespace GetRevisions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the specified document file
			document.LoadFromFile("..\..\..\..\..\..\..\Data\GetRevisions.docx")

			' Initialize a StringBuilder to store insert revisions
			Dim insertRevision As New StringBuilder()
			insertRevision.AppendLine("Insert revisions:")

			' Initialize an index for insert revisions
			Dim index_insertRevision As Integer = 0

			' Initialize a StringBuilder to store delete revisions
			Dim deleteRevision As New StringBuilder()
			deleteRevision.AppendLine("Delete revisions:")

			' Initialize an index for delete revisions
			Dim index_deleteRevision As Integer = 0

			' Iterate through each section in the document
			For Each sec As Section In document.Sections

				' Iterate through each DocumentObject in the section's body
				For Each docItem As DocumentObject In sec.Body.ChildObjects
					If TypeOf docItem Is Paragraph Then
						Dim para As Paragraph = CType(docItem, Paragraph)

						' Check if the paragraph contains insert revision
						If para.IsInsertRevision Then
							' Increment the insert revision index
							index_insertRevision += 1
							insertRevision.AppendLine("Index: " & index_insertRevision)

							' Get the InsertRevision object from the paragraph
							Dim insRevison As EditRevision = para.InsertRevision

							' Get the type of insert revision
							Dim insType As EditRevisionType = insRevison.Type
							insertRevision.AppendLine("Type: " & insType)

							' Get the author of the insert revision
							Dim insAuthor As String = insRevison.Author
							insertRevision.AppendLine("Author: " & insAuthor)

						ElseIf para.IsDeleteRevision Then
							' Increment the delete revision index
							index_deleteRevision += 1
							deleteRevision.AppendLine("Index: " & index_deleteRevision)

							' Get the DeleteRevision object from the paragraph
							Dim delRevison As EditRevision = para.DeleteRevision

							' Get the type of delete revision
							Dim delType As EditRevisionType = delRevison.Type
							deleteRevision.AppendLine("Type: " & delType)

							' Get the author of the delete revision
							Dim delAuthor As String = delRevison.Author
							deleteRevision.AppendLine("Author: " & delAuthor)
						End If

						' Iterate through each DocumentObject in the paragraph
						For Each obj As DocumentObject In para.ChildObjects
							If TypeOf obj Is TextRange Then
								Dim textRange As TextRange = CType(obj, TextRange)

								' Check if the TextRange contains insert revision
								If textRange.IsInsertRevision Then
									' Increment the insert revision index
									index_insertRevision += 1
									insertRevision.AppendLine("Index: " & index_insertRevision)

									' Get the InsertRevision object from the TextRange
									Dim insRevison As EditRevision = textRange.InsertRevision

									' Get the type of insert revision
									Dim insType As EditRevisionType = insRevison.Type
									insertRevision.AppendLine("Type: " & insType)

									' Get the author of the insert revision
									Dim insAuthor As String = insRevison.Author
									insertRevision.AppendLine("Author: " & insAuthor)
								ElseIf textRange.IsDeleteRevision Then
									' Increment the delete revision index
									index_deleteRevision += 1
									deleteRevision.AppendLine("Index: " & index_deleteRevision)

									' Get the DeleteRevision object from the TextRange
									Dim delRevison As EditRevision = textRange.DeleteRevision

									' Get the type of delete revision
									Dim delType As EditRevisionType = delRevison.Type
									deleteRevision.AppendLine("Type: " & delType)

									' Get the author of the delete revision
									Dim delAuthor As String = delRevison.Author
									deleteRevision.AppendLine("Author: " & delAuthor)
								End If
							End If
						Next obj
					End If
				Next docItem
			Next sec

			' Write the insert revisions to a text file
			File.WriteAllText("insertRevisions.txt", insertRevision.ToString())

			' Write the delete revisions to a text file
			File.WriteAllText("deleteRevisions.txt", deleteRevision.ToString())

			' Dispose of the document object
			document.Dispose()

		End Sub
	End Class
End Namespace
