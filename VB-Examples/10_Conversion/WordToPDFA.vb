Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace WordToPDFA
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document.
			Dim document As New Document()

			'Load the file from disk.
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			'Create a ToPdfParameterList instance
			Dim toPdf As New ToPdfParameterList()

			'Set the Conformance-level of the Pdf file to PDF_A1B.
			toPdf.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B

			Dim result As String = "Result-WordToPDFA.pdf"

			'Save the file.
			document.SaveToFile(result, toPdf)

			'Dispose the document.
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
