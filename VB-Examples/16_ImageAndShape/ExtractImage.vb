Imports System.Drawing.Imaging
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Interface

Namespace ExtractImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object
			Dim document As New Document("..\..\..\..\..\..\Data\Template.docx")

			'Create a queue to store composite objects
			Dim nodes As New Queue(Of ICompositeObject)()

			'Enqueue the document as the initial node
			nodes.Enqueue(document)

			'Create a list to store images
			Dim images As IList(Of Image) = New List(Of Image)()

			'Traverse through the composite objects in the document
			Do While nodes.Count > 0

				'Dequeue the next node
				Dim node As ICompositeObject = nodes.Dequeue()

				'Iterate through the child objects of the node
				For Each child As IDocumentObject In node.ChildObjects

					'If the child is a composite object, enqueue it for further processing
					If TypeOf child Is ICompositeObject Then
						nodes.Enqueue(TryCast(child, ICompositeObject))
						If child.DocumentObjectType = DocumentObjectType.Picture Then

							'If the child is a picture, add its image to the list
							Dim picture As DocPicture = TryCast(child, DocPicture)
							images.Add(picture.Image)
						End If
					End If
				Next child
			Loop

			'Save each image in the list as a PNG file
			For i As Integer = 0 To images.Count - 1
				Dim fileName As String = String.Format("Image-{0}.png", i)
				images(i).Save(fileName, ImageFormat.Png)
			Next i

			'Dispose the document
			document.Dispose()

			If images.Count > 0 Then
				'show the first image
				Process.Start("Image-0.png")
			End If
		End Sub

	End Class
End Namespace
