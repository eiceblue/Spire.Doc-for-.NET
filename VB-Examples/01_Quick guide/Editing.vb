Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace Editing
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
            Dim document_Renamed As New Document()

            'load a document
			document_Renamed.LoadFromFile("..\..\..\..\..\..\Data\Editing.doc")

			'Get a paragraph
			Dim paragraph_Renamed As Paragraph = document_Renamed.Sections(0).AddParagraph()

			'Append Text
			paragraph_Renamed.AppendText("Editing sample")

			'Save doc file.
			document_Renamed.SaveToFile("Sample.doc", FileFormat.Doc)

			'Launching the MS Word file.
			WordDocViewer("Sample.doc")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace