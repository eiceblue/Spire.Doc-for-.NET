Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetDocumentProperties
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			' Set the title, author, company, keywords, and comments for the document's built-in properties
			document.BuiltinDocumentProperties.Title = "Document Demo Document"
			document.BuiltinDocumentProperties.Author = "James"
			document.BuiltinDocumentProperties.Company = "e-iceblue"
			document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo"
			document.BuiltinDocumentProperties.Comments = "This document is just a demo."

			' Access the custom document properties and add new custom properties
			Dim custom As CustomDocumentProperties = document.CustomDocumentProperties
			custom.Add("e-iceblue", True)
			custom.Add("Authorized By", "John Smith")
			custom.Add("Authorized Date", Date.Today)

			' Save the modified document to a new file in Docx format
			document.SaveToFile("Output.docx", FileFormat.Docx)

			' Release all resources used by the Document object
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
