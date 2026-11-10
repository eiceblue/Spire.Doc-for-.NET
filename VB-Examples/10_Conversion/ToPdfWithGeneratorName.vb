Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc

Namespace ToPdfWithGeneratorName
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document instance
			Dim document As New Document()

			' Load the Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\..\Data\ConvertedTemplate.docx")

			' Create a ToPdfParameterList instance to configure PDF conversion options
			Dim toPdf As New ToPdfParameterList()

			' Define the generator name
			toPdf.GeneratorName = "Spire.Doc for .NET Product"
			document.SaveToFile("ToPdfWithGeneratorName.pdf", toPdf)
			document.Close()
			document.Dispose()

			'view the PDF file.
			WordDocViewer("ToPdfWithGeneratorName.pdf")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
