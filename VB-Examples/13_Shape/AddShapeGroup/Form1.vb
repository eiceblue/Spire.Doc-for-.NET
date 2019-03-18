Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddShapeGroup
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
				 'create a document
			Dim doc As New Document()
			Dim sec As Section = doc.AddSection()

			'add a new paragraph
			Dim para As Paragraph = sec.AddParagraph()
			'add a shape group with the height and width
			Dim shapegroup As ShapeGroup = para.AppendShapeGroup(375, 462)
			shapegroup.HorizontalPosition = 180
			'calcuate the scale ratio
			Dim X As Single = CSng(shapegroup.Width / 1000.0f)
			Dim Y As Single = CSng(shapegroup.Height / 1000.0f)

			Dim txtBox As New Spire.Doc.Fields.TextBox(doc)
			txtBox.SetShapeType(ShapeType.RoundRectangle)
			txtBox.Width = 125 / X
			txtBox.Height = 54 / Y
			Dim paragraph As Paragraph = txtBox.Body.AddParagraph()
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
			paragraph.AppendText("Start")
			txtBox.HorizontalPosition = 19/ X
			txtBox.VerticalPosition = 27 / Y
			txtBox.Format.LineColor = Color.Green
			shapegroup.ChildObjects.Add(txtBox)

			Dim arrowLineShape As New ShapeObject(doc, ShapeType.DownArrow)
			arrowLineShape.Width = 16 / X
			arrowLineShape.Height = 40 / Y
			arrowLineShape.HorizontalPosition = 69 / X
			arrowLineShape.VerticalPosition = 87 / Y
			arrowLineShape.StrokeColor = Color.Purple
			shapegroup.ChildObjects.Add(arrowLineShape)

			txtBox = New Spire.Doc.Fields.TextBox(doc)
			txtBox.SetShapeType(ShapeType.Rectangle)
			txtBox.Width = 125 / X
			txtBox.Height = 54 / Y
			paragraph = txtBox.Body.AddParagraph()
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
			paragraph.AppendText("Step 1")
			txtBox.HorizontalPosition = 19/ X
			txtBox.VerticalPosition = 131/ Y
			txtBox.Format.LineColor = Color.Blue
			shapegroup.ChildObjects.Add(txtBox)

			arrowLineShape = New ShapeObject(doc, ShapeType.DownArrow)
			arrowLineShape.Width = 16 / X
			arrowLineShape.Height = 40 / Y
			arrowLineShape.HorizontalPosition = 69 / X
			arrowLineShape.VerticalPosition = 192 / Y
			arrowLineShape.StrokeColor = Color.Purple
			shapegroup.ChildObjects.Add(arrowLineShape)

			txtBox = New Spire.Doc.Fields.TextBox(doc)
			txtBox.SetShapeType(ShapeType.Parallelogram)
			txtBox.Width = 149 / X
			txtBox.Height = 59/ Y
			paragraph = txtBox.Body.AddParagraph()
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
			paragraph.AppendText("Step 2")
			txtBox.HorizontalPosition = 7 / X
			txtBox.VerticalPosition = 236/ Y
			txtBox.Format.LineColor = Color.BlueViolet
			shapegroup.ChildObjects.Add(txtBox)

			arrowLineShape = New ShapeObject(doc, ShapeType.DownArrow)
			arrowLineShape.Width = 16 / X
			arrowLineShape.Height = 40/ Y
			arrowLineShape.HorizontalPosition = 66 / X
			arrowLineShape.VerticalPosition = 300 / Y
			arrowLineShape.StrokeColor = Color.Purple
			shapegroup.ChildObjects.Add(arrowLineShape)

			txtBox = New Spire.Doc.Fields.TextBox(doc)
			txtBox.SetShapeType(ShapeType.Rectangle)
			txtBox.Width = 125 / X
			txtBox.Height = 54 / Y
			paragraph = txtBox.Body.AddParagraph()
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
			paragraph.AppendText("Step 3")
			txtBox.HorizontalPosition = 19 / X
			txtBox.VerticalPosition = 345 / Y
			txtBox.Format.LineColor = Color.Blue
			shapegroup.ChildObjects.Add(txtBox)



			'save the document
			doc.SaveToFile("ShapeGroup.docx", FileFormat.Docx2010)

			FileViewer("ShapeGroup.docx")

		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
