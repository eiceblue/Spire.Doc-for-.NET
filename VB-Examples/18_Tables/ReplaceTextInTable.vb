Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ReplaceTextInTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
		   ' Create a new document object
		   Dim doc As New Document()

		   ' Load a document from a file, specified by the file path
		   doc.LoadFromFile("..\..\..\..\..\..\Data\ReplaceTextInTable.docx")

		   ' Get the first section of the document
		   Dim section As Section = doc.Sections(0)

		   ' Get the first table in the section
		   Dim table As Table = TryCast(section.Tables(0), Table)

		   ' Create a regular expression pattern for matching text within curly braces
		   Dim regex As New System.Text.RegularExpressions.Regex("{[^\}]+\}")

		   ' Replace text in the table that matches the regular expression pattern with "E-iceblue"
		   table.Replace(regex, "E-iceblue")

		   ' Replace the text "Beijing" with "Component" in the table, case-insensitive and match whole words only
		   table.Replace("Beijing", "Component", False, True)

		   ' Specify the output file name
		   Dim output As String = "ReplaceTextInTable_out.docx"

		   ' Save the modified document to a file, using Docx2013 format
		   doc.SaveToFile(output, FileFormat.Docx2013)

		   ' Dispose of the document object
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
