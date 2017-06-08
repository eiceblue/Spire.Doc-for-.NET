Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace Merge
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim fileName As String = OpenFile()
			Dim fileMerge As String = OpenFile()
			If ((Not String.IsNullOrEmpty(fileName))) AndAlso ((Not String.IsNullOrEmpty(fileMerge))) Then
				'Create word document
				Dim document_Renamed As New Document()
				document_Renamed.LoadFromFile(fileName,FileFormat.Doc)

				Dim documentMerge As New Document()
				documentMerge.LoadFromFile(fileMerge, FileFormat.Doc)

				For Each sec As Section In documentMerge.Sections
					document_Renamed.Sections.Add(sec.Clone())
				Next sec

				'Save doc file.
				document_Renamed.SaveToFile("Sample.doc", FileFormat.Doc)

				'Launching the MS Word file.
				WordDocViewer("Sample.doc")
			End If


		End Sub

		Private Function OpenFile() As String
			openFileDialog1.Filter = "Word Document (*.doc)|*.doc"
			openFileDialog1.Title = "Choose a document to merage"

			openFileDialog1.RestoreDirectory = True
            If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Return openFileDialog1.FileName
            End If

			Return String.Empty
		End Function

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
