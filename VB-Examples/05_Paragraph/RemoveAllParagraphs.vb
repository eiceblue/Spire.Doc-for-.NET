Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemoveAllParagraphs
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new instance of the Document class.
			Dim document As New Document()

			'Load a Word document from a specified file path using relative path.
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			'Clear all paragraphs in each section of the document.
			For Each section As Section In document.Sections
				section.Paragraphs.Clear()
			Next section

			'Set the file name for the resulting document after removing all paragraphs.
			Dim result As String = "Result-RemoveAllParagraphs.docx"

			'Save the modified document to the specified file path in the Docx2013 format.
			document.SaveToFile(result, FileFormat.Docx2013)

			'Dispose of the document object to release any associated resources.
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
