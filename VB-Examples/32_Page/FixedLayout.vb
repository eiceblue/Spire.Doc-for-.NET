Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Pages

Namespace FixedLayout
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim inputFile As String = "..\..\..\..\..\..\Data\Template_Docx_3.docx"

			' Create a new instance of Document
			Dim doc As New Document()

			' Load the document from the specified file
			doc.LoadFromFile(inputFile, FileFormat.Docx)

			' Create a FixedLayoutDocument from the loaded document
			Dim layoutDoc As New FixedLayoutDocument(doc)

			' Get the first line in the first column of the first page
			Dim line As FixedLayoutLine = layoutDoc.Pages(0).Columns(0).Lines(0)

			' Create a StringBuilder to store the output text
			Dim stringBuilder As New StringBuilder()
			stringBuilder.AppendLine("Line: " & line.Text)

			' Get the paragraph that contains the line and append its text to the StringBuilder
			Dim para As Paragraph = line.Paragraph
			stringBuilder.AppendLine("Paragraph text: " & para.Text)

			' Get the text content of the first page
			Dim pageText As String = layoutDoc.Pages(0).Text
			stringBuilder.AppendLine(pageText)

			' Iterate through each page in the FixedLayoutDocument
			For Each page As FixedLayoutPage In layoutDoc.Pages
				' Get all the lines on the current page
				Dim lines As LayoutCollection(Of LayoutElement) = page.GetChildEntities(LayoutElementType.Line, True)

				' Append the page index and number of lines to the StringBuilder
				stringBuilder.AppendLine("Page " & page.PageIndex & " has " & lines.Count & " lines.")
			Next page

			' Append the lines of the first paragraph to the StringBuilder
			' (except runs and nodes in the header and footer).
			stringBuilder.AppendLine("The lines of the first paragraph:")
			For Each paragraphLine As FixedLayoutLine In layoutDoc.GetLayoutEntitiesOfNode((CType(doc.FirstChild, Section)).Body.Paragraphs(0))
				stringBuilder.AppendLine(paragraphLine.Text.Trim())
				stringBuilder.AppendLine(paragraphLine.Rectangle.ToString())
			Next paragraphLine

			' Write the contents of the StringBuilder to a text file
			File.WriteAllText("page.txt", stringBuilder.ToString())

			' Dispose of the document object when finished using it
			doc.Dispose()		
			
		WordDocViewer("page.txt")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch e As Exception
				Debug.Write(e.StackTrace)
			End Try
		End Sub
	End Class
End Namespace
