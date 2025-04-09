Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports System.Text.RegularExpressions

Namespace ReplaceContentWithDoc
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object and load the first document from the specified file path
			Dim document1 As New Document()
			document1.LoadFromFile("..\..\..\..\..\..\Data\ReplaceContentWithDoc.docx")

			' Create a new Document object and load the second document from the specified file path
			Dim document2 As New Document()
			document2.LoadFromFile("..\..\..\..\..\..\Data\Insert.docx")

			' Get the first section of document1
			Dim section1 As Section = document1.Sections(0)

			' Create a regular expression pattern to match "[MY_DOCUMENT]"
			Dim regex As New Regex("\[MY_DOCUMENT\]", RegexOptions.None)

			' Find all text sections in document1 that match the pattern
			Dim textSections() As TextSelection = document1.FindAllPattern(regex)

			' Iterate through each TextSelection
			For Each selection As TextSelection In textSections
				' Get the owning Paragraph and TextRange of the selection
				Dim para As Paragraph = selection.GetAsOneRange().OwnerParagraph
				Dim textRange As TextRange = selection.GetAsOneRange()

				' Get the index of the owning Paragraph within section1
				Dim index As Integer = section1.Body.ChildObjects.IndexOf(para)

				' Iterate through each Section in document2
				For Each section2 As Section In document2.Sections
					' Iterate through each Paragraph in section2 and insert a clone into section1 at the specified index
					For Each paragraph As Paragraph In section2.Paragraphs
						section1.Body.ChildObjects.Insert(index, TryCast(paragraph.Clone(), Paragraph))
					Next paragraph
				Next section2

				' Remove the textRange from the owning Paragraph
				para.ChildObjects.Remove(textRange)
			Next selection

			' Save the modified document1 to a file named "Output.docx"
			document1.SaveToFile("Output.docx", FileFormat.Docx)

			' Dispose of the Document objects to release resources
			document1.Dispose()
			document2.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
