Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertPageBreakFirstApproach
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load an existing document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_2.docx")

			' Find all occurrences of the string "technology" in the document and store the results in an array of TextSelection objects
			Dim selections() As TextSelection = document.FindAllString("technology", True, True)

			' Iterate through each TextSelection object in the array
			For Each ts As TextSelection In selections
				' Get the entire range of text for the current TextSelection
				Dim range As TextRange = ts.GetAsOneRange()
				' Get the paragraph that contains the found text range
				Dim paragraph As Paragraph = range.OwnerParagraph
				' Get the index of the found text range within its parent paragraph
				Dim index As Integer = paragraph.ChildObjects.IndexOf(range)

				' Create a new page break and insert it after the found text range within the paragraph
				Dim pageBreak As New Break(document, BreakType.PageBreak)
				paragraph.ChildObjects.Insert(index + 1, pageBreak)
			Next ts

			' Specify the file name for the resulting document
			Dim result As String = "Result-InsertPageBreakAtSpecifiedLocation.docx"

			' Save the modified document to a new file with the specified file format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object
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
