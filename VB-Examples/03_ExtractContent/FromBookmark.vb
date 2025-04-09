Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace FromBookmark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object named sourcedocument
			Dim sourcedocument As New Document()

			' Load a Word document from a specific file path into sourcedocument
			sourcedocument.LoadFromFile("..\..\..\..\..\..\Data\Bookmark.docx")

			' Create a new Document object named destinationDoc
			Dim destinationDoc As New Document()

			' Add a section to the destinationDoc
			Dim section As Section = destinationDoc.AddSection()

			' Add a paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Create a BookmarksNavigator object with the sourcedocument
			Dim navigator As New BookmarksNavigator(sourcedocument)

			' Move the navigator to the bookmark named "Test"
			navigator.MoveToBookmark("Test", True, True)

			' Get the content of the bookmark as a TextBodyPart
			Dim textBodyPart As TextBodyPart = navigator.GetBookmarkContent()

			' Create a list to store TextRanges
			Dim list As New List(Of TextRange)()

			' Iterate over the BodyItems in the textBodyPart
			For Each item In textBodyPart.BodyItems

				' Check if the item is a Paragraph
				If TypeOf item Is Paragraph Then

					' Iterate over the ChildObjects in the Paragraph
					For Each childObject In (TryCast(item, Paragraph)).ChildObjects

						' Check if the childObject is a TextRange
						If TypeOf childObject Is TextRange Then

							' Cast childObject to TextRange and add it to the list
							Dim range As TextRange = TryCast(childObject, TextRange)
							list.Add(range)
						End If
					Next childObject
				End If
			Next item

			' Iterate over the list of TextRanges
			For m As Integer = 0 To list.Count - 1
				
				' Clone each TextRange and add it to the paragraph
				paragraph.Items.Add(list(m).Clone())
			Next m

			' Save the destinationDoc to a file named "Output.docx" in DOCX format
			destinationDoc.SaveToFile("Output.docx", FileFormat.Docx)

			' Dispose the sourcedocument object to release its resources
			sourcedocument.Dispose()

			' Dispose the destinationDoc object to release its resources
			destinationDoc.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
