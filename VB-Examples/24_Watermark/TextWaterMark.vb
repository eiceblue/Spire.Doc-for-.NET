Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace TextWaterMark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a template file
			Dim document As New Document("..\..\..\..\..\..\Data\Template.docx")

			' Insert text watermark into the first section of the document
			InsertTextWatermark(document.Sections(0))

			' Save the modified document to a new file
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Dispose the document object
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("Sample.docx")


		End Sub
		Private Shared Sub InsertTextWatermark(ByVal section As Section)
            ' Create a TextWatermark object
            Dim txtWatermark As New Spire.Doc.TextWatermark()
            ' Set the text for the watermark
            txtWatermark.Text = "E-iceblue"
			' Set the font size of the watermark
			txtWatermark.FontSize = 95
			' Set the color of the watermark
			txtWatermark.Color = Color.Blue
			' Set the layout of the watermark
			txtWatermark.Layout = WatermarkLayout.Diagonal
			' Set the watermark for the document section
			section.Document.Watermark = txtWatermark
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
