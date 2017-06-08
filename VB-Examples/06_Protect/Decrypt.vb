Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace Decrypt
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim fileName As String = OpenFile()
			If Not String.IsNullOrEmpty(fileName) Then
				'Create word document
				Dim document_Renamed As New Document()
				document_Renamed.LoadFromFile(fileName,FileFormat.Doc,Me.textBox1.Text)

				'Save doc file.
				document_Renamed.SaveToFile("Sample.doc", FileFormat.Doc)

				'Launching the MS Word file.
				WordDocViewer("Sample.doc")
			End If


		End Sub

        Private Function OpenFile() As String
            openFileDialog1.InitialDirectory _
                = New System.IO.DirectoryInfo("..\..\..\..\..\..\Data").FullName
            openFileDialog1.FileName = "Protect_Decrypt.doc"
            openFileDialog1.Filter = "Word Document (*.doc)|*.doc"
            openFileDialog1.Title = "Choose a document to Decrypt"

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
