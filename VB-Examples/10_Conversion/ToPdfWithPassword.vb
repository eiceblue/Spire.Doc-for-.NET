Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc

Namespace ToPdfWithPassword
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load a Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\..\Data\ConvertedTemplate.docx")

			' Create a new ToPdfParameterList object to set parameters for PDF conversion
			Dim toPdf As New ToPdfParameterList()

			' Specify the password for encrypting the PDF
			Dim password As String = "E-iceblue"

			' Set PDF security options to encrypt the PDF with the specified password and permissions
			toPdf.PdfSecurity.Encrypt(password, password, Spire.Doc.PdfPermissionsFlags.Default, Spire.Doc.PdfEncryptionKeySize.Key128Bit)

			' Save the document to a PDF file with the specified encryption settings
			document.SaveToFile("EncryptWithPassword.pdf", toPdf)

			' Dispose of the Document object to release resources
			document.Dispose()


			'view the PDF file.
			WordDocViewer("EncryptWithPassword.pdf")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
