Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc.Fields
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace InsertFootnote
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim document As New Document()

			' Load the Word document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\FootnoteExample.docx")

			' Find the specified string in the document
			Dim selection As TextSelection = document.FindString("Spire.Doc", False, True)

			' Get the selected text as a single range
			Dim textRange As TextRange = selection.GetAsOneRange()

			' Get the paragraph that contains the selected text
			Dim paragraph As Paragraph = textRange.OwnerParagraph

			' Get the index of the selected text within the paragraph's child objects
			Dim index As Integer = paragraph.ChildObjects.IndexOf(textRange)

			' Append a footnote to the paragraph
			Dim footnote As Footnote = paragraph.AppendFootnote(FootnoteType.Footnote)

			' Insert the footnote into the paragraph's child objects at the specified index
			paragraph.ChildObjects.Insert(index + 1, footnote)

			' Add a paragraph to the footnote's text body and append text to it
			textRange = footnote.TextBody.AddParagraph().AppendText("Welcome to evaluate Spire.Doc")

			' Set the font name, size, and color for the appended text
			textRange.CharacterFormat.FontName = "Arial Black"
			textRange.CharacterFormat.FontSize = 10
			textRange.CharacterFormat.TextColor = Color.DarkGray

			' Set the font name, size, style, and color for the footnote marker
			footnote.MarkerCharacterFormat.FontName = "Calibri"
			footnote.MarkerCharacterFormat.FontSize = 12
			footnote.MarkerCharacterFormat.Bold = True
			footnote.MarkerCharacterFormat.TextColor = Color.DarkGreen

			' Save the modified document to a file
			document.SaveToFile("AddFootnote.docx", FileFormat.Docx2010)

			' Dispose of the document object when finished using it
			document.Dispose()

			'view the Word file.
			WordDocViewer("AddFootnote.docx")
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
