Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemoveSmartArt
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Const SmartArtDefaultWidth As Integer = 432
		Private Const SmartArtDefaultHeight As Integer = 252
		Private Const TitleFontSize As Single = 28f
		Private Const DefaultNodeFontSize As Single = 15f
		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class to represent a Word document
			Dim document As New Document()

			' Load an existing Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\SmartArt.docx")

			' Iterate through each paragraph in the first section of the document
			For j As Integer = 0 To document.Sections(0).Paragraphs.Count - 1
				' Get the current paragraph from the first section
				Dim paragraph As Paragraph = document.Sections(0).Paragraphs(j)

				' Iterate through each child object within the current paragraph
				For i As Integer = 0 To paragraph.ChildObjects.Count - 1
					' Check if the current child object is a Shape (which can contain SmartArt)
					If TypeOf paragraph.ChildObjects(i) Is Spire.Doc.Fields.Shapes.Shape Then
						' Cast the child object to a Shape object
						Dim shape As Spire.Doc.Fields.Shapes.Shape = TryCast(paragraph.ChildObjects(i), Spire.Doc.Fields.Shapes.Shape)

						' Check if this shape contains a SmartArt graphic
						If shape.HasSmartArt Then
							' Remove the SmartArt shape from the paragraph's items collection
							paragraph.Items.RemoveAt(i)

							' Decrement the loop counter since we removed an item and the next item has shifted down
							i -= 1
						End If
					End If
				Next i
			Next j

			' Define the output file name for saving the modified document
			Dim result As String = "RemoveSmartArt.docx"

			' Save the modified document (with SmartArt removed) to a file in Docx2016 format
			document.SaveToFile(result, FileFormat.Docx2016)

			' Close the document to release any file handles or resources
			document.Close()

			' Dispose of the document object to free up system memory
			document.Dispose()

			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
