Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace Replace
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

			'Replace text
			document_Renamed.Replace(Me.textBox1.Text, Me.textBox2.Text,True,True)

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
