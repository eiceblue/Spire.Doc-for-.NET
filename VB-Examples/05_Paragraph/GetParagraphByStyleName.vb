Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.IO

Namespace GetParagraphByStyleName
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_DocX_3.docx")

			' Create a StringBuilder object to store the content
			Dim content As New StringBuilder()

			' Append a line of text to describe the operation
			content.AppendLine("Get paragraphs by style name ""Heading1"": ")

			' Iterate through each Section in the document
			For Each section As Section In document.Sections

				' Iterate through each Paragraph in the current Section
				For Each paragraph As Paragraph In section.Paragraphs
				
					' Check if the current Paragraph has the style name "Heading1"
					If paragraph.StyleName = "Heading1" Then
					
						' Append the text of the matching Paragraph to the content
						content.AppendLine(paragraph.Text)
						
					End If
					
				Next paragraph
				
			Next section

			' Specify the filename for the output result
			Dim result As String = "Result-GetParagraphsByStyleName.txt"

			' Write the content to the specified file
			File.WriteAllText(result, content.ToString())

			' Dispose the Document object to release resources
			document.Dispose()

			'Launch the file.
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
