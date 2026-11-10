Imports Spire.Doc
Imports Spire.Doc.Documents.Comparison

Namespace CompareDocumentsWithOptions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object (doc1)
			Dim doc1 As New Document()

			' Load a document from a file ("..\..\..\..\..\..\Data\SupportDocumentCompare1.docx")
			doc1.LoadFromFile("..\..\..\..\..\..\Data\SupportDocumentCompare1.docx")

			' Create a new Document object (doc2)
			Dim doc2 As New Document()

			' Load another document from a file ("..\..\..\..\..\..\Data\SupportDocumentCompare2.docx")
			doc2.LoadFromFile("..\..\..\..\..\..\Data\SupportDocumentCompare2.docx")

			' Create a new CompareOptions object
			Dim compareOptions As New CompareOptions()

			' Set the IgnoreFormatting property to True to ignore formatting during comparison
			compareOptions.IgnoreFormatting = True

			' Compare doc1 with doc2 using "E-iceblue" as the author and the current date
			doc1.Compare(doc2, "E-iceblue", Date.Now, compareOptions)

			' Define the file name for the resulting document
			Dim result As String = "CompareDocumentsWithOptions_result.docx"

			' Save doc1 to a file using the specified file format (Docx2013)
			doc1.SaveToFile(result, Spire.Doc.FileFormat.Docx2013)

			' Dispose of the Document objects to release resources
			doc1.Dispose()
			doc2.Dispose()
			
			'View the document
			FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
