Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertWordArt
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create Word document.
			Dim doc As New Document()

			'Load Word document.
			doc.LoadFromFile("..\..\..\..\..\..\Data\InsertWordArt.docx")

			'Add a paragraph.
			Dim paragraph As Paragraph = doc.Sections(0).AddParagraph()

			'Add a shape.
			Dim shape As ShapeObject = paragraph.AppendShape(250, 70, ShapeType.TextWave4)

			'Set the position of the shape.
			shape.VerticalPosition = 20
			shape.HorizontalPosition = 80

			'Set the text of WordArt.
			shape.WordArt.Text = "Thanks for reading."

			'Set the fill color.
			shape.FillColor = Color.Red

			'Set the border color of the text.
			shape.StrokeColor = Color.Yellow

			'Save docx file.
			doc.SaveToFile("WordArt.docx", FileFormat.Docx2013)
			
			'Dispose the document
			doc.Dispose()

			'Launch the Word file.
			FileViewer("WordArt.docx")
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
