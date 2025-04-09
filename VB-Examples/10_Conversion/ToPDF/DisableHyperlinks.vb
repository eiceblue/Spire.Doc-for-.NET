Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace DisableHyperlinks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\..\Data\Template_Docx_5.docx")

			' Create a new ToPdfParameterList object to customize PDF conversion settings
			Dim pdf As New ToPdfParameterList()

			'Set DisableLink to true to remove the hyperlink effect for the result PDF page. 
			'Set DisableLink to false to preserve the hyperlink effect for the result PDF page.
			pdf.DisableLink = True

			' Define the output file name for the PDF file with customized settings
			Dim result As String = "Result-DisableHyperlinks.pdf"
			document.SaveToFile(result, pdf)

			' Dispose of the Document object to release resources
			document.Dispose()

			'Launch the Pdf file.
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
