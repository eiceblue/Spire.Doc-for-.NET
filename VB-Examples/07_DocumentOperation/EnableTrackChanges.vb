Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace EnableTrackChanges
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object.
			Dim document As New Document()

			'Load a document from a specified file path.
			document.LoadFromFile("..\..\..\..\..\..\..\Data\Template_Docx_1.docx")

			'Enable track changes in the document.
			document.TrackChanges = True

			'Specify the file name for the resulting document.
			Dim result As String = "Result-EnableTrackChanges.docx"

			'Save the document to the specified file path using Docx2013 format.
			document.SaveToFile(result, FileFormat.Docx2013)

			'Dispose of the document object to free up resources.
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
