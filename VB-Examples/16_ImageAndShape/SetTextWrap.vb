Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetTextWrap
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

			'Loop through the sections
			For Each sec As Section In doc.Sections

				'Loop through the paragraphs
				For Each para As Paragraph In sec.Paragraphs

					'Create a list to store the pictures
					Dim pictures As New List(Of DocumentObject)()

					'Get all pictures in the Word document
					For Each docObj As DocumentObject In para.ChildObjects
						If docObj.DocumentObjectType = DocumentObjectType.Picture Then
							pictures.Add(docObj)
						End If
					Next docObj

					'Set text wrap styles for each piture
					For Each pic As DocumentObject In pictures

						'Create a DocPicture instance
						Dim picture As DocPicture = TryCast(pic, DocPicture)
						picture.TextWrappingStyle = TextWrappingStyle.Through
						picture.TextWrappingType = TextWrappingType.Both
					Next pic
				Next para
			Next sec

			'Save the document
			Dim output As String = "SetTextWrap.docx"
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
