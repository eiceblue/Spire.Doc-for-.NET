Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace WordToPdfEncrypt
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the Word document file from the specified path
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_2.docx")

			' Create a ToPdfParameterList object to specify PDF conversion parameters
			Dim toPdf As New ToPdfParameterList()

			' Encrypt the PDF with the specified password "e-iceblue"
			toPdf.PdfSecurity.Encrypt("e-iceblue")

			' Specify the output file name for the converted PDF
			Dim result As String = "Result-WordToPdfEncrypt.pdf"

			' Save the document as a PDF with the specified encryption settings
			document.SaveToFile(result, toPdf)

			' Dispose the Document object to free resources
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
