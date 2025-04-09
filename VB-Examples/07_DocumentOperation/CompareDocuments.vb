Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace CompareDocuments
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object for doc1
			Dim doc1 As New Document()

			'Load the file "SupportDocumentCompare1.docx" into doc1
			doc1.LoadFromFile("..\..\..\..\..\..\Data\SupportDocumentCompare1.docx")

			'Create a new Document object for doc2
			Dim doc2 As New Document()

			'Load the file "SupportDocumentCompare2.docx" into doc2
			doc2.LoadFromFile("..\..\..\..\..\..\Data\SupportDocumentCompare2.docx")

			'Compare doc1 with doc2 using the "E-iceblue" comparison option
			doc1.Compare(doc2, "E-iceblue")

			'Specify the file name for the result document as "CompareDocuments_result.docx"
			Dim result As String = "CompareDocuments_result.docx"

			'Save doc1 to the specified file name in Docx2013 format
			doc1.SaveToFile(result, FileFormat.Docx2013)

			'Dispose of the doc1 object to release resources
			doc1.Dispose()

			'Dispose of the doc2 object to release resources
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
