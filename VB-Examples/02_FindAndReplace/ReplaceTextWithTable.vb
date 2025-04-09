Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ReplaceTextWithTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object.
			Dim document As New Document()

			' Load the Word document from the specified file path.
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			' Get the first section of the document.
			Dim section As Section = document.Sections(0)

			' Find the first occurrence of the text "Christmas Day, December 25" in the document and retrieve it as a TextSelection object.
			Dim selection As TextSelection = document.FindString("Christmas Day, December 25", True, True)

			' Convert the selected text into a single TextRange.
			Dim range As TextRange = selection.GetAsOneRange()

			' Get the paragraph that owns the text range.
			Dim paragraph As Paragraph = range.OwnerParagraph

			' Get the body that owns the paragraph.
			Dim body As Body = paragraph.OwnerTextBody

			' Retrieve the index of the paragraph within its owning body.
			Dim index As Integer = body.ChildObjects.IndexOf(paragraph)

			' Add a table to the section.
			Dim table As Table = section.AddTable(True)

			' Reset the cells in the table to have 3 rows and 3 columns.
			table.ResetCells(3, 3)

			' Remove the paragraph from its owning body.
			body.ChildObjects.Remove(paragraph)

			' Insert the table at the original index of the paragraph within the body.
			body.ChildObjects.Insert(index, table)

			' Specify the output file name for saving the modified document.
			Dim result As String = "Result-ReplaceTextWithTable.docx"

			' Save the modified document to the specified output file in Docx2013 format.
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the Document object to release resources.
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
