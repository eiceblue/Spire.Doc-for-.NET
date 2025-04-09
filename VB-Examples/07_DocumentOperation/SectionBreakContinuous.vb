Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.Text

Namespace SectionBreakContinuous
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new Document object and load a Word document from the specified file path
			Dim doc As New Document("..\..\..\..\..\..\..\Data\Sample_two sections.docx")

			' Iterate through each Section in the document
			For Each section As Section In doc.Sections
				' Set the BreakCode property of each section to "NoBreak"
				section.BreakCode = SectionBreakType.NoBreak
			Next section

			' Save the modified document to a new file with the specified file format
			doc.SaveToFile("result.docx")

			' Dispose of the resources used by the Document object
			doc.Dispose()

			'Launch result file
			WordDocViewer("result.docx")

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
