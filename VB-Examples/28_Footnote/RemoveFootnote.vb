Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace RemoveFootnote
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim document As New Document()

			' Load the Word document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\Footnote.docx")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Iterate through each paragraph in the section
			For Each para As Paragraph In section.Paragraphs
				Dim index As Integer = -1

				' Find the index of the first footnote within the paragraph's child objects
				Dim i As Integer = 0
				Dim cnt As Integer = para.ChildObjects.Count
				Do While i < cnt
					Dim pBase As ParagraphBase = TryCast(para.ChildObjects(i), ParagraphBase)

					If TypeOf pBase Is Footnote Then
						index = i
						Exit Do
					End If
					i += 1
				Loop

				' If a footnote is found, remove it from the paragraph's child objects
				If index > -1 Then
					para.ChildObjects.RemoveAt(index)
				End If
			Next para

			' Save the modified document to a file
			document.SaveToFile("RemoveFootnote.docx", FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'view the Word file.
			WordDocViewer("RemoveFootnote.docx")
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
