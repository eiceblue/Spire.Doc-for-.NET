Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ConvertFieldToBodyText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object to store the source document
			Dim sourceDocument As New Document()

			' Load the source document from a file
			sourceDocument.LoadFromFile("..\..\..\..\..\..\Data\TextInputField.docx")

			' Iterate through each form field in the first section of the document's body
			For Each field As FormField In sourceDocument.Sections(0).Body.FormFields
				' Check if the form field is of type FieldFormTextInput
				If field.Type = FieldType.FieldFormTextInput Then
					' Get the owner paragraph of the form field
					Dim paragraph As Paragraph = field.OwnerParagraph

					' Initialize variables for start and end index of bookmark objects
					Dim startIndex As Integer = 0
					Dim endIndex As Integer = 0

					' Create a TextRange object using the source document
					Dim textRange As New TextRange(sourceDocument)

					' Set the text of the TextRange to the text of the paragraph
					textRange.Text = paragraph.Text

					' Iterate through each child object of the paragraph
					For Each obj As DocumentObject In paragraph.ChildObjects
						' Check if the child object is a BookmarkStart object
						If obj.DocumentObjectType = DocumentObjectType.BookmarkStart Then
							' Store the index of the BookmarkStart object
							startIndex = paragraph.ChildObjects.IndexOf(obj)
						End If

						' Check if the child object is a BookmarkEnd object
						If obj.DocumentObjectType = DocumentObjectType.BookmarkEnd Then
							' Store the index of the BookmarkEnd object
							endIndex = paragraph.ChildObjects.IndexOf(obj)
						End If
					Next obj

					' Remove the form fields or child objects between the start and end index
					For i As Integer = endIndex To startIndex + 1 Step -1
						If TypeOf paragraph.ChildObjects(i) Is TextFormField Then
							' Remove the TextFormField object
							Dim textFormField As TextFormField = TryCast(paragraph.ChildObjects(i), TextFormField)
							paragraph.ChildObjects.Remove(textFormField)
						Else
							' Remove other child objects
							paragraph.ChildObjects.RemoveAt(i)
						End If
					Next i

					' Insert the modified TextRange at the start index of the paragraph
					paragraph.ChildObjects.Insert(startIndex, textRange)

					' Exit the loop after processing the first FieldFormTextInput
					Exit For
				End If
			Next field

			' Save the modified document to a new file
			sourceDocument.SaveToFile("Output.docx", FileFormat.Docx)

			' Dispose the source document object
			sourceDocument.Dispose()

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
