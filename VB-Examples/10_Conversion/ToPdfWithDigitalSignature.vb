Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc

Namespace ToPdfWithDigitalSignature
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object to represent a Word document.
			Dim document As New Document()

			' Load the Word document from the specified relative file path.
			document.LoadFromFile("..\..\..\..\..\..\..\Data\Sample.docx")

			' Create a parameter list object to configure PDF conversion settings.
			Dim toPdfParameter As New ToPdfParameterList()

			' Configure digital signature information using a PFX certificate file, password, reason, and location.
			toPdfParameter.DigitalSignatureInfo = New PdfDigitalSignatureInfo("..\..\..\..\..\..\..\Data\gary.pfx", "e-iceblue", "signing", "Spire Doc")

			' Set the hash algorithm used for the digital signature to RIPEMD-160.
			toPdfParameter.DigitalSignatureInfo.HashAlgorithm = PdfDigitalSignatureHashAlgorithm.RipeMD160

			' Set the signature timestamp to the current date and time.
			toPdfParameter.DigitalSignatureInfo.SignatureDate = Date.Now

			' Configure the timestamp server (TSA) settings for the digital signature, including URL, username, and password.
			toPdfParameter.DigitalSignatureInfo.TimestampConfiguration = New PdfDigitalSignatureTimestampConfiguration("https://freetsa.org/tsr", "Gary", "Password")

			' Save the document as a PDF file with the specified digital signature settings.
			document.SaveToFile("ToPdfWithDigitalSignature.pdf", toPdfParameter)

			' Close the document to release internal resources.
			document.Close()

			' Explicitly dispose of the document object to free memory immediately.
			document.Dispose()

			'view the PDF file.
			WordDocViewer("ToPdfWithDigitalSignature.pdf")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
