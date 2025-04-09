Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace StartFromFormField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object to store the source document.
			Dim sourceDocument As New Document()

			' Load the source document from a file.
			sourceDocument.LoadFromFile("..\..\..\..\..\..\Data\TextInputField.docx")

			' Create a new Document object to store the destination document.
			Dim destinationDoc As New Document()

			' Add a new section to the destination document.
			Dim section As Section = destinationDoc.AddSection()

			' Initialize an index variable.
			Dim index As Integer = 0

			' Iterate through each form field in the first section of the source document.
			For Each field As FormField In sourceDocument.Sections(0).Body.FormFields

				' Check if the form field is a text input field.
				If field.Type = FieldType.FieldFormTextInput Then

					' Get the paragraph that contains the form field.
					Dim paragraph As Paragraph = field.OwnerParagraph

					' Determine the index of the paragraph in the source document.
					index = sourceDocument.Sections(0).Body.ChildObjects.IndexOf(paragraph)

					' Exit the loop after finding the first text input field.
					Exit For
				End If
			Next field

			' Copy the next three child objects starting from the index found above to the destination document's section.
			For i As Integer = index To index + 3 - 1

				' Clone the document object at the specified index.
				Dim dobj As DocumentObject = sourceDocument.Sections(0).Body.ChildObjects(i).Clone()

				' Add the cloned document object to the body of the destination document's section.
				section.Body.ChildObjects.Add(dobj)
			Next i

			' Save the destination document to a file named "FromFormField.docx" in Docx format.
			destinationDoc.SaveToFile("FromFormField.docx", FileFormat.Docx)

			' Dispose of the source document to release resources.
			sourceDocument.Dispose()

			' Dispose of the destination document to release resources.
			destinationDoc.Dispose()

			'Launch the Word file.
			WordDocViewer("FromFormField.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
