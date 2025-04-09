Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ExtractParagraphBasedOnStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Define the style name for heading 1
			Dim styleName1 As String = "Heading1"

			' Create a StringBuilder object to store the text with style 1
			Dim style1Text As New StringBuilder()

			' Load the document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\ExtractParagraphBasedOnStyle.docx")

			' Append a header for the text with style 1
			style1Text.AppendLine("The following is the content of the paragraph with the style name " & styleName1 & ": ")

			' Iterate over sections in the document
			For Each section As Section In document.Sections

				' Iterate over paragraphs in each section
				For Each paragraph As Paragraph In section.Paragraphs

					' Check if the paragraph has the desired style name
					If paragraph.StyleName IsNot Nothing AndAlso paragraph.StyleName.Equals(styleName1) Then
						
						' Append the text of the paragraph to the style1Text
						style1Text.AppendLine(paragraph.Text)
					End If
				Next paragraph
			Next section

			' Specify the output file name
			Dim output1 As String = "ExtractParagraphBasedOnStyle_style1.txt"

			' Write the contents of style1Text to the output file
			File.WriteAllText(output1, style1Text.ToString())

			' Dispose the document object
			document.Dispose()

			WordDocViewer(output1)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
