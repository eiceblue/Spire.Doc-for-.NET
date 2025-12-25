Imports Spire.Doc
Imports Spire.Doc.Documents.Comparison

Namespace CompareDocumentsIgnoreTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the first document from the specified file path
			Dim document1 As Document = New Document("..\..\..\..\..\..\Data\ComparedDoc1.docx")

			' Load the second document from the specified file path
			Dim document2 As Document = New Document("..\..\..\..\..\..\Data\ComparedDoc2.docx")

			' Create a new CompareOptions object to specify comparison settings
			Dim compareoptions As CompareOptions = New CompareOptions()

			' Set the option to ignore differences in tables during comparison
			compareoptions.IgnoreTable = True

			' Compare the two documents using the specified options, with "E-iceblue" as the author name for tracked changes
			document1.Compare(document2, "E-iceblue", compareoptions)

			' Save the compared document (with changes tracked) to a new file in DOCX 2019 format
			document1.SaveToFile("CompareDocumentsIgnoreTable.docx", FileFormat.Docx2019)

			' Release resources used by the first document
			document1.Dispose()

			' Release resources used by the second document
			document2.Dispose()

			' Launching the MS Word file.
			WordDocViewer("CompareDocumentsIgnoreTable.docx")

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
