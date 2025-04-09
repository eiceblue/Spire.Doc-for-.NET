Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Fields
Imports Spire.Doc.Documents
Namespace HandleAskField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document
			Dim doc As New Document()

			' Load the document from a file
			doc.LoadFromFile("..\..\..\..\..\..\Data\HandleAskField.docx")

			' Subscribe to the UpdateFields event
			AddHandler doc.UpdateFields, AddressOf doc_UpdateFields

			' Enable field update
			doc.IsUpdateFields = True

			' Save the modified document to a file
			doc.SaveToFile("HandleAskField.docx", FileFormat.Docx)

			' Dispose the document object
			doc.Dispose()
			
			WordDocViewer("HandleAskField.docx")

		End Sub
		
		' Event handler for updating fields
		Private Shared Sub doc_UpdateFields(ByVal sender As Object, ByVal args As IFieldsEventArgs)
			' Check if the event arguments are of type AskFieldEventArgs
			If TypeOf args Is AskFieldEventArgs Then
				Dim askArgs As AskFieldEventArgs = TryCast(args, AskFieldEventArgs)

				' Handle different bookmarks and set response text accordingly
				If askArgs.BookmarkName = "1" Then
					askArgs.ResponseText = "Thank you. This is my first time to come to a Chinese restaurant. Could you tell me the different features of Chinese food?"
				End If

				If askArgs.BookmarkName = "2" Then
					askArgs.ResponseText = "No more, thank you. I'm quite full."
				End If
			End If
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
