Imports Spire.Doc
Imports Spire.Doc.Collections
Imports Spire.Doc.Documents
Imports Spire.Doc.Interface

Namespace DeletePictureBullet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Initialize a new Document object to work with Word documents.
			Dim doc As New Document()

			' Load an existing Word document from the specified file path into the document object.
			doc.LoadFromFile("..\..\..\..\..\..\Data\PictureBullets.docx")

			' Access the collection of list definitions (styles) within the document.
			Dim lists As ListCollection = doc.ListReferences

			' Get the second level (index 1) of the first list definition (index 0) in the document.
			Dim listLevel As ListLevel = lists(0).Levels(1)

			' Remove any picture bullet associated with this list level.
			listLevel.DeletePictureBullet()

			' Save the modified document to a new file named "DeletePictureBullet.docx" in Word 2016 format.

			doc.SaveToFile("DeletePictureBullet.docx", FileFormat.Docx2016)
			' Close the document to release internal resources such as file handles.

			doc.Close()

			' Explicitly dispose of the document object to ensure memory is freed immediately.
			doc.Dispose()

			'Launching the Word file.
			WordDocViewer("DeletePictureBullet.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
