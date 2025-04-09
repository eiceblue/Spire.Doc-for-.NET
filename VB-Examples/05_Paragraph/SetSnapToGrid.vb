Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetSnapToGrid
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim doc As New Document()

			' Add a new section to the document
			Dim section As Section = doc.AddSection()

			' Set the grid type of the page to display only lines
			section.PageSetup.GridType = GridPitchType.LinesOnly

			' Set the number of lines per page to 15
			section.PageSetup.LinesPerPage = 15

			' Add a new paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Add text content to the paragraph
			paragraph.AppendText("With Spire.Doc, you can generate, modify, convert, render and print documents without utilizing Microsoft Word®. But you need MS Word viewer to view the resultant document.")

			' Enable snapping to the grid for the paragraph formatting
			paragraph.Format.SnapToGrid = True

			' Specify the output file name for the modified document
			Dim output As String = "SetSnapToGrid.docx"

			' Save the modified document to the specified file format (Docx2013)
			doc.SaveToFile(output, FileFormat.Docx2013)

			' Dispose of the document object to release resources
			doc.Dispose()

			'Launch the file 
			WordDocViewer(output)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace

