Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace UpdateImage
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

			'Create a list to store the pictures
			Dim pictures As New List(Of DocumentObject)()

			'Loop through the sections
			For Each sec As Section In doc.Sections

				'Loop through the paragraphs
				For Each para As Paragraph In sec.Paragraphs

					'Loop through the child objects of the paragraph
					For Each docObj As DocumentObject In para.ChildObjects

						'Determine if the type is picture or not
						If docObj.DocumentObjectType = DocumentObjectType.Picture Then

							'Add the picure to list
							pictures.Add(docObj)
						End If
					Next docObj
				Next para
			Next sec

			'Create a DocPicture instance
			Dim picture As DocPicture = TryCast(pictures(0), DocPicture)

			'Replace the first picture with a new image file
			picture.LoadImage(Image.FromFile("..\..\..\..\..\..\Data\E-iceblue.png"))

			'Save the document
			Dim output As String = "ReplaceWithNewImage.docx"
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
