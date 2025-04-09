Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ToSVG
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document
			Dim document As New Document()

			'Load the file from disk
			document.LoadFromFile("..\..\..\..\..\..\Data\ToSVGTemplate.docx")

			'Save to file
			document.SaveToFile("Sample.svg", FileFormat.SVG)

			'Dispose the document
			document.Dispose()

			'Launching the svg file.
			Process.Start("Sample.svg")
		End Sub
	End Class
End Namespace
