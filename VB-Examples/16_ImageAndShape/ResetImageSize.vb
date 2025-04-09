Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ResetImageSize
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\ImageTemplate.docx"

			'Create a word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first secion
			Dim section As Section = doc.Sections(0)

			'Get the first paragraph
			Dim paragraph As Paragraph = section.Paragraphs(0)

			'Reset the image size of the first paragraph
			For Each docObj As DocumentObject In paragraph.ChildObjects
				If TypeOf docObj Is DocPicture Then
					Dim picture As DocPicture = TryCast(docObj, DocPicture)

					'Set the width
					picture.Width = 50.0F

					'Set the height
					picture.Height = 50.0F
				End If
			Next docObj

			'Save the document
			Dim output As String = "ResetImageSize.docx"
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
