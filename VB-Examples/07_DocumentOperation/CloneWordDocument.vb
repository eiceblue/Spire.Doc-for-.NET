Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace CloneWordDocument
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object
			Dim document As New Document()

			'Load a Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\..\Data\Template_Docx_1.docx")

			'Clone the document and assign it to a new Document object
			Dim newDoc As Document = document.Clone()

			'Save the cloned document to a new file with the specified name and format (Docx2013)
			Dim result As String = "Result-CloneWordDocument.docx"
			newDoc.SaveToFile(result, FileFormat.Docx2013)

			'Dispose of the original document object to free up resources
			document.Dispose()

			'Dispose of the cloned document object to free up resources
			newDoc.Dispose()

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
