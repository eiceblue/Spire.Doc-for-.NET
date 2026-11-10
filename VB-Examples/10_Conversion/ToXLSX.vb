Imports Spire.Doc

Namespace ToXLSX
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load an existing Word document
			document.LoadFromFile("..\..\..\..\..\..\Data\ConvertedToXLSX.docx")

			' Define the file path and name for the output document
			Dim result As String = "ToXLSX.xlsx"

			' Convert the Word document to XLSX file
			document.SaveToFile(result, FileFormat.XLSX)

			' Close the document to release resources
			document.Close()

			' Dispose of the document object to free up memory
			document.Dispose()

			'Launching the Word file.
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
