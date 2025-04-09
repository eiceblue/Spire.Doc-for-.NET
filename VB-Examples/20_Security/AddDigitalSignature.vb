Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Namespace AddDigitalSignature
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
		
			' Create a new Document object
			Dim doc As New Document()

			' Load the document from the specified file path
			doc.LoadFromFile("..\..\..\..\..\..\Data\AddDigitalSignature.doc")

			' Specify the output file path for the signed document with digital signature
			Dim result As String = "AddDigitalSignature_result.docx"

			' Save the document to the output file path in DOCX format with the specified certificate and password
			doc.SaveToFile(result, FileFormat.Docx, "..\..\..\..\..\..\Data\gary.pfx", "e-iceblue")

			' Dispose the document object to free up resources
			doc.Dispose()
			
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
