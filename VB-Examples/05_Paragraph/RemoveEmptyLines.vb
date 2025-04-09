Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemoveEmptyLines
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object.
			Dim document As New Document()

			' Load a Word document from the specified file path.
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_3.docx")

			' Iterate through each section in the document.
			For Each section As Section In document.Sections
				' Initialize a counter variable.
				Dim i As Integer = 0

				' Loop through the child objects in the section's body.
				Do While i < section.Body.ChildObjects.Count
					' Check if the child object is a paragraph.
					If section.Body.ChildObjects(i).DocumentObjectType = DocumentObjectType.Paragraph Then
						' Convert the child object to a Paragraph and check if the text is empty or consists only of whitespace.
						If String.IsNullOrEmpty((TryCast(section.Body.ChildObjects(i), Paragraph)).Text.Trim()) Then
							' Remove the empty paragraph from the section's body.
							section.Body.ChildObjects.Remove(section.Body.ChildObjects(i))
							' Decrement the counter to adjust for the removed object.
							i -= 1
						End If
					End If

					' Increment the counter to move to the next child object.
					i += 1
				Loop
			Next section

			' Specify the file name for the resulting document after removing empty lines.
			Dim result As String = "Result-RemoveEmptyLines.docx"

			' Save the modified document to the specified file path in Docx2013 format.
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the Document object to release resources.
			document.Dispose()

			'Launch the MS Word file.
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
