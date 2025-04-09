Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AddPageBorders
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the specified document file
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			' Access the first section of the document
			Dim section As Section = document.Sections(0)

			' Set the border type for the page setup of the section
			section.PageSetup.Borders.BorderType = Spire.Doc.Documents.BorderStyle.DoubleWave

			' Set the color for the page borders
			section.PageSetup.Borders.Color = Color.LightSeaGreen

			' Set the space for the left border of the page setup
			section.PageSetup.Borders.Left.Space = 50

			' Set the space for the right border of the page setup
			section.PageSetup.Borders.Right.Space = 50

			' Specify the file path for the output result
			Dim result As String = "Result-AddPageBorders.docx"

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
