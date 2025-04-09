Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace RotateShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\Shapes.docx"

			'Create a Word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first section
			Dim section As Section = doc.Sections(0)

			'Traverse the word document and set the shape rotation as 20
			For Each para As Paragraph In section.Paragraphs
				For Each obj As DocumentObject In para.ChildObjects
					If TypeOf obj Is ShapeObject Then
						'Set the shape rotation as 20
						TryCast(obj, ShapeObject).Rotation = 20.0
					End If
				Next obj
			Next para

			'Save and launch document
			Dim output As String = "RotateShape.docx"
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
