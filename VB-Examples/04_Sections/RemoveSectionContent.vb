Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace RemoveSectionContent
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object called doc
			Dim doc As New Document()

			'Load a Word document from the specified file path
			doc.LoadFromFile("..\..\..\..\..\..\Data\Template_N3.docx")

			'Iterate through each Section in the document
			For Each section As Section In doc.Sections

				'Clear the child objects in the header of the current section
				section.HeadersFooters.Header.ChildObjects.Clear()

				'Clear the child objects in the body of the current section
				section.Body.ChildObjects.Clear()

				'Clear the child objects in the footer of the current section
				section.HeadersFooters.Footer.ChildObjects.Clear()
			Next section

			'Specify the output file name
			Dim output As String = "RemoveSectionContent_out.docx"

			'Save the modified document as a Word document with the specified file format
			doc.SaveToFile(output, FileFormat.Docx2013)

			'Dispose of the doc object to release resources
			doc.Dispose()

			'Launch the file
			FileViewer(output)
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
