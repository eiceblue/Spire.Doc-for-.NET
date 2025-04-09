Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace Macros
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim document As New Document()

			' Load the Word document from the specified file that may contain VBA macros
			document.LoadFromFile("../../../../../../Data/Macros.docm", FileFormat.Docm)

			' Save the document to a new file with the specified name and format (Docm for macro-enabled document)
			document.SaveToFile("Sample.docm", FileFormat.Docm)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("Sample.docm")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
