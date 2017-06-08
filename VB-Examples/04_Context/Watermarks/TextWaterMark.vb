Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace TextWaterMark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Open a blank word document as template
			Dim document As New Document("..\..\..\..\..\..\..\Data\Blank.doc")
			InsertTextWatermark(document.Sections(0))
			'Save doc file.
			document.SaveToFile("Sample.doc",FileFormat.Doc)

			'Launching the MS Word file.
			WordDocViewer("Sample.doc")


		End Sub
        Private Sub InsertTextWatermark(ByVal section As Section)
            Dim paragraph As Paragraph
            If (section.Paragraphs.Count > 0) Then
                paragraph = section.Paragraphs(0)
            Else
                paragraph = section.AddParagraph()
            End If
            paragraph.AppendText("The sample demonstrates how to insert text watermark into a document.")
            paragraph.ApplyStyle(BuiltinStyle.Heading2)


            Dim txtWatermark As New Spire.Doc.TextWatermark()
            txtWatermark.Text = "Watermark Demo"
            txtWatermark.FontSize = 90
            txtWatermark.Color = Color.Red
            txtWatermark.Layout = WatermarkLayout.Diagonal
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
