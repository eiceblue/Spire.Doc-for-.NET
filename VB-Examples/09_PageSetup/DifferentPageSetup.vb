Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.Text

Namespace DifferentPageSetup
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new instance of the Document class with the specified file path
			Dim doc As New Document("..\..\..\..\..\..\Data\DifferentPageSetup.docx")

			' Access the second section of the document
			Dim SectionTwo As Section = doc.Sections(1)

			' Set the page orientation of the second section to landscape
			SectionTwo.PageSetup.Orientation = PageOrientation.Landscape

			' Set the page size of the second section (uncomment this line if needed)
			'SectionTwo.PageSetup.PageSize = new SizeF(800, 800);

			' Save the modified document to a new file
			doc.SaveToFile("result.docx")

			' Dispose of the document object
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
