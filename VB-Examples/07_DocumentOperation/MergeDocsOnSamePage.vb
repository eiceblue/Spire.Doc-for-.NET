Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace MergeDocsOnSamePage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()
			' Load a Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\..\Data\Insert.docx")

			' Create a new Document object for the destination document
			Dim destinationDocument As New Document()
			' Load another Word document as the destination document from the specified file path
			destinationDocument.LoadFromFile("..\..\..\..\..\..\..\Data\TableOfContent.docx")

			' Iterate through each Section in the source document
			For Each section As Section In document.Sections
				' Iterate through each DocumentObject in the body of the current Section
				For Each obj As DocumentObject In section.Body.ChildObjects
					' Add a clone of the DocumentObject to the body of the first Section in the destination document
					destinationDocument.Sections(0).Body.ChildObjects.Add(obj.Clone())
				Next obj
			Next section

			' Save the destination document to a new file with the specified file format
			destinationDocument.SaveToFile("Output.docx", FileFormat.Docx)

			' Dispose of the resources used by the Document objects
			document.Dispose()
			destinationDocument.Dispose()

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
