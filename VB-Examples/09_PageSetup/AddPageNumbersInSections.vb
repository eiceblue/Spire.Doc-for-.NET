Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AddPageNumbersInSections
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the specified document file
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_4.docx")

			' Iterate through the first 3 sections of the document
			For i As Integer = 0 To 2
				' Access the footer of the current section and add a new paragraph
				Dim footer As HeaderFooter = document.Sections(i).HeadersFooters.Footer
				Dim footerParagraph As Paragraph = footer.AddParagraph()
				
				' Append page number field and "of" text to the footer paragraph
				footerParagraph.AppendField("page number", FieldType.FieldPage)
				footerParagraph.AppendText(" of ")
				footerParagraph.AppendField("number of pages", FieldType.FieldSectionPages)
				
				' Set the horizontal alignment of the footer paragraph to right
				footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right

				' If it's not the last section, restart page numbering and set the starting number for the next section
				If i = 2 Then
					Exit For
				Else
					document.Sections(i + 1).PageSetup.RestartPageNumbering = True
					document.Sections(i + 1).PageSetup.PageStartingNumber = 1
				End If
			Next i

			' Specify the file path for the output result
			Dim result As String = "Result-AddPageNumbersInSections.docx"

			' Save the modified document to a new file with the specified file format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object
			document.Dispose()

			'Launch the Ms Word file.
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
