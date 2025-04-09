Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ResetShapeSize
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\Shapes.docx"

			'Create a word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first section
			Dim section As Section = doc.Sections(0)

			'Get the first paragraph
			Dim para As Paragraph = section.Paragraphs(0)

			'Get the second shape
			Dim shape As ShapeObject = TryCast(para.ChildObjects(1), ShapeObject)

			'Reset the width and height of the shape
			shape.Width = 200
			shape.Height = 200

			'Save the document
			Dim output As String = "ResetShapeSize.docx"
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
