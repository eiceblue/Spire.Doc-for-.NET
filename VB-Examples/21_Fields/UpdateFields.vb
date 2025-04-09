Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections

Namespace UpdateFields
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a file
			Dim document As New Document("..\..\..\..\..\..\Data\IfFieldSample.docx")

			' Setting the culture source when updating fields
			document.FieldOptions.CultureSource = Spire.Doc.Layout.Fields.FieldCultureSource.CurrentThread

			' Enable automatic update of fields in the document
			document.IsUpdateFields = True

			' Save the document to a new file
			document.SaveToFile("result.docx", FileFormat.Docx)

			' Dispose of the document object
			document.Dispose()

			'Launch the Word file
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
