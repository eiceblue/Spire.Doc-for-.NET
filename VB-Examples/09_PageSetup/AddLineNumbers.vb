Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AddLineNumbers
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

			' Set the start value for line numbering in the first section's page setup
			document.Sections(0).PageSetup.LineNumberingStartValue = 1

			' Set the step value for line numbering in the first section's page setup
			document.Sections(0).PageSetup.LineNumberingStep = 6

			' Set the distance from text for line numbering in the first section's page setup
			document.Sections(0).PageSetup.LineNumberingDistanceFromText = 40.0F

			' Set the restart mode for line numbering in the first section's page setup
			document.Sections(0).PageSetup.LineNumberingRestartMode = LineNumberingRestartMode.Continuous

			' Specify the file path for the output result
			Dim result As String = "Result-AddLineNumbers.docx"

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
