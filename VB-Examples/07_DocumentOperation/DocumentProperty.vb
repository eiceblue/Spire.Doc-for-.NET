Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace DocumentProperty
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object and load the specified file
			Dim document As New Document("..\..\..\..\..\..\Data\Summary_of_Science.doc")

			' Set the title property of the document
			document.BuiltinDocumentProperties.Title = "Document Demo Document"

			' Set the subject property of the document
			document.BuiltinDocumentProperties.Subject = "demo"

			' Set the author property of the document
			document.BuiltinDocumentProperties.Author = "James"

			' Set the company property of the document
			document.BuiltinDocumentProperties.Company = "e-iceblue"

			' Set the manager property of the document
			document.BuiltinDocumentProperties.Manager = "Jakson"

			' Set the category property of the document
			document.BuiltinDocumentProperties.Category = "Doc Demos"

			' Set the keywords property of the document
			document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo"

			' Set the comments property of the document
			document.BuiltinDocumentProperties.Comments = "This document is just a demo."

			' Save the modified document to a new file in Docx format
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Release all resources used by the Document object
			document.Dispose()


			'Launching the MS Word file.
			WordDocViewer("Sample.docx")


		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
