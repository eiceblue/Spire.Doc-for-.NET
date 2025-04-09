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
		   'Create a new document
			Dim doc As New Document()

			'Add a section to the document
			Dim sec As Section = doc.AddSection()

			'Add a new paragraph
			Dim para As Paragraph = sec.AddParagraph()

			'Create a shape group and set its width and height
			Dim shapegroup As ShapeGroup = para.AppendShapeGroup(375, 462)

			'Set the horizontal position of the shape group
			shapegroup.HorizontalPosition = 180

			'Calculate the scaling factors for width and height
			Dim X As Single = CSng(shapegroup.Width / 1000.0F)
			Dim Y As Single = CSng(shapegroup.Height / 1000.0F)
			Dim txtBox As New Spire.Doc.Fields.TextBox(doc)
			txtBox.SetShapeType(ShapeType.RoundRectangle)
			txtBox.Width = 125 / X
			txtBox.Height = 54 / Y
			Dim paragraph As Paragraph = txtBox.Body.AddParagraph()
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
			paragraph.AppendText("Start")
			txtBox.HorizontalPosition = 19 / X
			txtBox.VerticalPosition = 27 / Y
			txtBox.Format.LineColor = Color.Green
			shapegroup.ChildObjects.Add(txtBox)

			'Add an arrow shape to the shape group
			Dim arrowLineShape As New ShapeObject(doc, ShapeType.DownArrow)
			arrowLineShape.Width = 16 / X
			arrowLineShape.Height = 40 / Y
			arrowLineShape.HorizontalPosition = 69 / X
			arrowLineShape.VerticalPosition = 87 / Y
			arrowLineShape.StrokeColor = Color.Purple
			shapegroup.ChildObjects.Add(arrowLineShape)

			'Add another text box to the shape group
			txtBox = New Spire.Doc.Fields.TextBox(doc)
			txtBox.SetShapeType(ShapeType.Rectangle)
			txtBox.Width = 125 / X
			txtBox.Height = 54 / Y
			paragraph = txtBox.Body.AddParagraph()
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
			paragraph.AppendText("Step 1")
			txtBox.HorizontalPosition = 19 / X
			txtBox.VerticalPosition = 131 / Y
			txtBox.Format.LineColor = Color.Blue
			shapegroup.ChildObjects.Add(txtBox)

			'Add another arrow shape to the shape group
			arrowLineShape = New ShapeObject(doc, ShapeType.DownArrow)
			arrowLineShape.Width = 16 / X
			arrowLineShape.Height = 40 / Y
			arrowLineShape.HorizontalPosition = 69 / X
			arrowLineShape.VerticalPosition = 192 / Y
			arrowLineShape.StrokeColor = Color.Purple
			shapegroup.ChildObjects.Add(arrowLineShape)

			'Add another text box to the shape group
			txtBox = New Spire.Doc.Fields.TextBox(doc)
			txtBox.SetShapeType(ShapeType.Parallelogram)
			txtBox.Width = 149 / X
			txtBox.Height = 59 / Y
			paragraph = txtBox.Body.AddParagraph()
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
			paragraph.AppendText("Step 2")
			txtBox.HorizontalPosition = 7 / X
			txtBox.VerticalPosition = 236 / Y
			txtBox.Format.LineColor = Color.BlueViolet
			shapegroup.ChildObjects.Add(txtBox)

			'Add another arrow shape to the shape group
			arrowLineShape = New ShapeObject(doc, ShapeType.DownArrow)
			arrowLineShape.Width = 16 / X
			arrowLineShape.Height = 40 / Y
			arrowLineShape.HorizontalPosition = 66 / X
			arrowLineShape.VerticalPosition = 300 / Y
			arrowLineShape.StrokeColor = Color.Purple
			shapegroup.ChildObjects.Add(arrowLineShape)

			'Add another text box to the shape group
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

			'Save the document
			doc.SaveToFile("ShapeGroup.docx", FileFormat.Docx2010)

			'Dispose the document
			doc.Dispose()

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
