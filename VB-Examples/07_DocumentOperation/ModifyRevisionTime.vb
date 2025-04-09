Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting.Revisions
Imports Spire.Doc.Fields
Imports System.IO

Namespace ModifyRevisionTime
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the specified document file
			document.LoadFromFile("..\..\..\..\..\..\..\Data\ModifyRevisionTime.docx")

			' Initialize index for insert revisions
			Dim index_insertRevision As Integer = 0

			' Initialize index for delete revisions
			Dim index_deleteRevision As Integer = 0

			' Specify the date string
			Dim dateString As String = "2023/3/1 00:00:00"

			' Specify the format for parsing the date string
			Dim format As String = "yyyy/M/d HH:mm:ss"

			' Parse the date string into a Date object
			Dim [date] As Date = Date.ParseExact(dateString, format, Nothing)

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
							' Get the InsertRevision object from the paragraph and set its DateTime property
							Dim insRevison As EditRevision = para.InsertRevision
							insRevison.DateTime = [date]
						ElseIf para.IsDeleteRevision Then
							' Increment the delete revision index
							index_deleteRevision += 1
							' Get the DeleteRevision object from the paragraph and set its DateTime property
							Dim delRevison As EditRevision = para.DeleteRevision
							delRevison.DateTime = [date]
						End If
						' Iterate through each DocumentObject in the paragraph
						For Each obj As DocumentObject In para.ChildObjects
							If TypeOf obj Is TextRange Then
								Dim textRange As TextRange = CType(obj, TextRange)
								' Check if the TextRange contains insert revision
								If textRange.IsInsertRevision Then
									' Increment the insert revision index
									index_insertRevision += 1
									' Get the InsertRevision object from the TextRange and set its DateTime property
									Dim insRevison As EditRevision = textRange.InsertRevision
									insRevison.DateTime = [date]
								ElseIf textRange.IsDeleteRevision Then
									' Increment the delete revision index
									index_deleteRevision += 1
									' Get the DeleteRevision object from the TextRange and set its DateTime property
									Dim delRevison As EditRevision = textRange.DeleteRevision
									delRevison.DateTime = [date]
								End If
							End If
						Next obj
					End If
				Next docItem
			Next sec

			' Save the modified document to a new file
			document.SaveToFile("ModifyRevisionTime_out.docx", FileFormat.Docx)

			' Dispose of the document object
			document.Dispose()

			WordDocViewer("ModifyRevisionTime_out.docx")

		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
