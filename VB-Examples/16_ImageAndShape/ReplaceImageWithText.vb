Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ReplaceImageWithText
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

			'Replace all pictures with texts
			Dim j As Integer = 1
			For Each sec As Section In doc.Sections
				For Each para As Paragraph In sec.Paragraphs
					Dim pictures As New List(Of DocumentObject)()
					'Get all pictures in the Word document
					For Each docObj As DocumentObject In para.ChildObjects
						If docObj.DocumentObjectType = DocumentObjectType.Picture Then
							pictures.Add(docObj)
						End If
					Next docObj

					'Replace pitures with the text "Here was image {image index}"
					For Each pic As DocumentObject In pictures

						'Get the index of the picture
						Dim index As Integer = para.ChildObjects.IndexOf(pic)

						'Create a new TextRange
						Dim range As New TextRange(doc)

						'Format the text
						range.Text = String.Format("Here was image {0}", j)

						'Insert the textrange
						para.ChildObjects.Insert(index, range)

						'Remove the picture
						para.ChildObjects.Remove(pic)
						j += 1
					Next pic
				Next para
			Next sec

			'Save the document
			Dim output As String = "ReplaceWithTexts.docx"
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
