Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace FindAndHighlight
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
            Dim document_Renamed As New Document()

            'load a document
			document_Renamed.LoadFromFile("..\..\..\..\..\..\Data\FindAndReplace.doc")

			'Find text
			Dim textSelections() As TextSelection = document_Renamed.FindAllString(Me.textBox1.Text, True, True)

			'Set hightlight
			For Each selection As TextSelection In textSelections
				selection.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow
			Next selection

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
