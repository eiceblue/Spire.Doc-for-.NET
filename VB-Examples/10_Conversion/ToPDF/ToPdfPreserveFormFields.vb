Imports Spire.Doc

Namespace ToPdfPreserveFormFields
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a file
			Dim document As New Document("..\..\..\..\..\..\..\Data\ToPdfPreserveFormFields.docx")

			' Preserve form field when converting to Pdf
			Dim ppl As New ToPdfParameterList()
			ppl.PreserveFormFields = True

			document.SaveToFile("ToPdfPreserveFormFields_output.pdf",ppl)
			' Dispose the document object
			document.Dispose()

			'Launch result file
			WordDocViewer("ToPdfPreserveFormFields_output.pdf")

		End Sub


		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
