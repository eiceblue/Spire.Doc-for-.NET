Imports Spire.Doc
Imports Spire.Doc.Collections
Imports Spire.Doc.Documents
Imports System.ComponentModel
Imports System.Text

Namespace FormACatalogue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object.
			Dim document As New Document()

			'Add a new section. 
			Dim section As Section = document.AddSection()
			Dim paragraph As Spire.Doc.Documents.Paragraph = If(section.Paragraphs.Count > 0, section.Paragraphs(0), section.AddParagraph())

			'Add Heading 1.
			paragraph = section.AddParagraph()
			paragraph.AppendText(BuiltinStyle.Heading1.ToString())
			paragraph.ApplyStyle(BuiltinStyle.Heading1)
			paragraph.ListFormat.ApplyNumberedStyle()

			'Add Heading 2.
			paragraph = section.AddParagraph()
			paragraph.AppendText(BuiltinStyle.Heading2.ToString())
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			'List style for Headings 2.

			Dim listStyle2 As ListStyle = document.Styles.Add(ListType.Numbered, "MyStyle2")
			Dim Levels As ListLevelCollection = listStyle2.ListRef.Levels
			For Each listLev As ListLevel In Levels
				listLev.UsePrevLevelPattern = True
				listLev.NumberPrefix = "1."
			Next listLev

			paragraph.ListFormat.ApplyStyle(listStyle2.Name)

			'Add list style 3.

			Dim listStyle3 As ListStyle = document.Styles.Add(ListType.Numbered, "MyStyle3")
			Dim Levels1 As ListLevelCollection = listStyle3.ListRef.Levels
			For Each listLev As ListLevel In Levels1
				listLev.UsePrevLevelPattern = True
				listLev.NumberPrefix = "1.1."
			Next listLev

			'Add Heading 3.
			For i As Integer = 0 To 3
				paragraph = section.AddParagraph()

				'Append text
				paragraph.AppendText(BuiltinStyle.Heading3.ToString())

				'Apply list style 3 for Heading 3
				paragraph.ApplyStyle(BuiltinStyle.Heading3)

				paragraph.ListFormat.ApplyStyle(listStyle3.Name)

			Next i

			' Specify the file name for the resulting Word document.
			Dim result As String = "Result-FormACatalogue.docx"

			' Save the Document object to a file in Docx format and dispose it.
			document.SaveToFile(result, FileFormat.Docx)
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
