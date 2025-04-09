Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddShapes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create Word document.
			Dim doc As New Document()

			'Add a section to the document
			Dim sec As Section = doc.AddSection()

			'Add a paragraph to the section
			Dim para As Paragraph = sec.AddParagraph()
			Dim x As Integer = 60, y As Integer = 40, lineCount As Integer = 0
			For i As Integer = 1 To 19

				'Check if the current line count is a multiple of 8
				If lineCount > 0 AndAlso lineCount Mod 8 = 0 Then

					'Append a page break to start a new page
					para.AppendBreak(BreakType.PageBreak)
					x = 60
					y = 40
					lineCount = 0
				End If

				'Append a shape to the paragraph
				Dim shape As ShapeObject = para.AppendShape(50, 50, CType(i, ShapeType))
				shape.HorizontalOrigin = HorizontalOrigin.Page
				shape.HorizontalPosition = x
				shape.VerticalOrigin = VerticalOrigin.Page
				shape.VerticalPosition = y + 50
				x = x + CInt(shape.Width) + 50

				'Check if the shape count is a multiple of 5
				If i > 0 AndAlso i Mod 5 = 0 Then

					'Adjust the vertical position and line count
					y = y + CInt(shape.Height) + 120
					lineCount += 1
					x = 60
				End If

			Next i

			'Save the document
			doc.SaveToFile("AddShape.docx", FileFormat.Docx)

			'Dispose the document
			doc.Dispose()

			'Launch Word file.
			WordDocViewer("AddShape.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace