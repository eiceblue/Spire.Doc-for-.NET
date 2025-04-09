Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace FormACatalogue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim document As New Document()

			' Add a section to the document
			Dim section As Section = document.AddSection()

			' Create a paragraph and add it to the section
			Dim paragraph As Paragraph = If(section.Paragraphs.Count > 0, section.Paragraphs(0), section.AddParagraph())

			' Add a new paragraph to the section
			paragraph = section.AddParagraph()

			' Set the text content of the paragraph as Heading1 style
			paragraph.AppendText(BuiltinStyle.Heading1.ToString())

			' Apply the Heading1 style to the paragraph
			paragraph.ApplyStyle(BuiltinStyle.Heading1)

			' Apply a numbered list format to the paragraph
			paragraph.ListFormat.ApplyNumberedStyle()

			' Add another paragraph to the section
			paragraph = section.AddParagraph()

			' Set the text content of the paragraph as Heading2 style
			paragraph.AppendText(BuiltinStyle.Heading2.ToString())

			' Apply the Heading2 style to the paragraph
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			' Create a new numbered list style with the ListType.Numbered type
			Dim listSty2 As New ListStyle(document, ListType.Numbered)

			' Configure the list levels of the list style
			For Each listLev As ListLevel In listSty2.Levels
				listLev.UsePrevLevelPattern = True
				listLev.NumberPrefix = "1."
			Next listLev

			' Set the name of the list style
			listSty2.Name = "MyStyle2"

			' Add the list style to the document's list styles collection
			document.ListStyles.Add(listSty2)

			' Apply the list style to the paragraph
			paragraph.ListFormat.ApplyStyle(listSty2.Name)

			' Create another numbered list style with the ListType.Numbered type
			Dim listSty3 As New ListStyle(document, ListType.Numbered)

			' Configure the list levels of the list style
			For Each listLev As ListLevel In listSty3.Levels
				listLev.UsePrevLevelPattern = True
				listLev.NumberPrefix = "1.1."
			Next listLev

			' Set the name of the list style
			listSty3.Name = "MyStyle3"

			' Add the list style to the document's list styles collection
			document.ListStyles.Add(listSty3)

			' Add paragraphs with Heading3 style and apply the list style
			For i As Integer = 0 To 3
				paragraph = section.AddParagraph()
				paragraph.AppendText(BuiltinStyle.Heading3.ToString())
				paragraph.ApplyStyle(BuiltinStyle.Heading3)
				paragraph.ListFormat.ApplyStyle(listSty3.Name)
			Next i

			' Specify the file name for the resulting document
			Dim result As String = "Result-FormACatalogue.docx"

			' Save the document to a file in Docx format
			document.SaveToFile(result, FileFormat.Docx)

			' Dispose the document object to free resources
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
