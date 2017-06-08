Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Forms
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Interface

Namespace Traversal
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            'open document
            Dim document As New Document("..\..\..\..\..\..\Data\Summary_of_Science.doc")

            'document elements, each of them has child elements
            Dim nodes As New Queue(Of ICompositeObject)()
            nodes.Enqueue(document)

            'embedded images list.
            Dim images As IList(Of Image) = New List(Of Image)()

            'traverse
            While nodes.Count > 0
                Dim node As ICompositeObject = nodes.Dequeue()
                For Each child As IDocumentObject In node.ChildObjects
                    If TypeOf child Is ICompositeObject Then
                        nodes.Enqueue(TryCast(child, ICompositeObject))
                    ElseIf child.DocumentObjectType = DocumentObjectType.Picture Then
                        Dim picture As DocPicture = TryCast(child, DocPicture)
                        images.Add(picture.Image)
                    End If
                Next
            End While

            'save images
            For i As Integer = 0 To images.Count - 1
                Dim fileName As [String] = [String].Format("Image-{0}.png", i)
                images(i).Save(fileName, ImageFormat.Png)
            Next

            If images.Count > 0 Then
                'show the first image
                System.Diagnostics.Process.Start("Image-0.png")
            End If

        End Sub
	End Class
End Namespace
