Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetLineShapeStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			'Create a document
			Dim doc As New Document()

			'Add a section
			Dim sec As Section = doc.AddSection()

			'Add a new paragraph
			Dim para As Paragraph = sec.AddParagraph()

			'Add a line shape
			Dim shape As ShapeObject = para.AppendShape(100, 100, ShapeType.Line)
			'Set style of Line shape
			shape.FillColor = Color.Orange
			shape.StrokeColor = Color.Black
			shape.LineStyle = ShapeLineStyle.Single
			shape.LineDashing = LineDashing.LongDashDotDotGEL

			'Save the document
			Dim output As String = "SetLineShapeStyle.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
			doc.Dispose()
			
			Viewer(output)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
