Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace KeepHiddenText
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

			' Create a new ToPdfParameterList object
			Dim pdf As New ToPdfParameterList()

			' Set the IsHidden property to True to save hidden text in the PDF
			pdf.IsHidden = True

			' Specify the file path for the resulting PDF
			Dim result As String = "Result-SaveTheHiddenTextToPDF.pdf"

			' Save the document to a PDF file with the specified parameters
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
