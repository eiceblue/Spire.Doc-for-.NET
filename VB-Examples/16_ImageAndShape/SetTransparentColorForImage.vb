Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetTransparentColorForImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\ImageTemplate.docx"

			'Create a Word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first paragraph in the first section
			Dim paragraph As Paragraph = doc.Sections(0).Paragraphs(0)

			'Loop through the child objects of the paragraph
			For Each obj As DocumentObject In paragraph.ChildObjects
				If TypeOf obj Is DocPicture Then
					'Set the blue color of the image(s) in the paragraph to transperant
					Dim picture As DocPicture = TryCast(obj, DocPicture)
					picture.TransparentColor = Color.Blue
				End If
			Next obj

			'Save the document
			Dim output As String = "SetTransparentColorForImage.docx"
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
