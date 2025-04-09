Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Fields
Imports Spire.Doc.Documents

Namespace InsertEndnote
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim doc As New Document()

			' Load a document from the specified file path
			doc.LoadFromFile("..\..\..\..\..\..\Data\InsertEndnote.doc")

			' Get the first section in the document
			Dim s As Section = doc.Sections(0)

			' Get the second paragraph in the section (index 1)
			Dim p As Paragraph = s.Paragraphs(1)

			' Append an endnote to the paragraph
			Dim endnote As Footnote = p.AppendFootnote(FootnoteType.Endnote)

			' Add a paragraph to the endnote's text body and append the reference text
			Dim text As TextRange = endnote.TextBody.AddParagraph().AppendText("Reference: Wikipedia")

			' Set the font name, size, and text color of the reference text
			text.CharacterFormat.FontName = "Impact"
			text.CharacterFormat.FontSize = 14
			text.CharacterFormat.TextColor = Color.DarkOrange

			' Set the font name, size, and text color of the endnote marker
			endnote.MarkerCharacterFormat.FontName = "Calibri"
			endnote.MarkerCharacterFormat.FontSize = 25
			endnote.MarkerCharacterFormat.TextColor = Color.DarkBlue

			' Save the modified document to the output file in DOCX format
			doc.SaveToFile("InsertEndnote.docx", FileFormat.Docx)

			' Dispose the document object
			doc.Dispose()

			'Launch the Word file
			WordDocViewer("InsertEndnote.docx")

		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
