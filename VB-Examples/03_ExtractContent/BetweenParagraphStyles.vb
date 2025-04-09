Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace BetweenParagraphStyles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object called sourceDocument.
			Dim sourceDocument As New Document()
			' Load a Word document from a specified file path.
			sourceDocument.LoadFromFile("..\..\..\..\..\..\Data\BetweenParagraphStyle.docx")

			' Create a new Document object called destinationDoc.
			Dim destinationDoc As New Document()

			' Add a Section to the destinationDoc.
			Dim section As Section = destinationDoc.AddSection()

			' Call the ExtractBetweenParagraphStyles method, passing the sourceDocument, destinationDoc, "1", and "2" as parameters.
			ExtractBetweenParagraphStyles(sourceDocument, destinationDoc, "1", "2")

			' Save the destinationDoc as a Word document with the filename "Output.docx".
			destinationDoc.SaveToFile("Output.docx", FileFormat.Docx)
			
			' Dispose of the source and destination Document objects to release resources.
			sourceDocument.Dispose()
			destinationDoc.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub

		' Define a private method called ExtractBetweenParagraphStyles, which takes sourceDocument, destinationDocument, stylename1, and stylename2 as parameters.
		Private Sub ExtractBetweenParagraphStyles(ByVal sourceDocument As Document, ByVal destinationDocument As Document, ByVal stylename1 As String, ByVal stylename2 As String)
			' Initialize startindex and endindex variables to 0.
			Dim startindex As Integer = 0
			Dim endindex As Integer = 0

			' Iterate through each Section in the sourceDocument.
			For Each section As Section In sourceDocument.Sections
				' Iterate through each Paragraph in the current section.
				For Each paragraph As Paragraph In section.Paragraphs
					' Check if the current paragraph's style name matches stylename1.
					If paragraph.StyleName = stylename1 Then
						' Set the startindex variable to the index of the current paragraph within the section.
						startindex = section.Body.Paragraphs.IndexOf(paragraph)
					End If
					' Check if the current paragraph's style name matches stylename2.
					If paragraph.StyleName = stylename2 Then
						' Set the endindex variable to the index of the current paragraph within the section.
						endindex = section.Body.Paragraphs.IndexOf(paragraph)
					End If
				Next paragraph

				' Iterate through each index between startindex and endindex (excluding both start and end).
				For i As Integer = startindex + 1 To endindex - 1
					' Clone the DocumentObject at the specified index from the sourceDocument and assign it to the doobj variable.
					Dim doobj As DocumentObject = sourceDocument.Sections(0).Body.ChildObjects(i).Clone()
					' Add the cloned DocumentObject to the child objects of the Section in the destinationDocument.
					destinationDocument.Sections(0).Body.ChildObjects.Add(doobj)
				Next i
			Next section
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
