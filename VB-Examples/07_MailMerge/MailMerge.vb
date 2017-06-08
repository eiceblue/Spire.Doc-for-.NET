Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace MailMerage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
			Dim document_Renamed As New Document()
			document_Renamed.LoadFromFile("..\..\..\..\..\..\Data\Fax.doc")

			Dim filedNames() As String = {"Contact Name","Fax","Date"}

			Dim filedValues() As String = {"John Smith","+1 (69) 123456",Date.Now.Date.ToString()}

			document_Renamed.MailMerge.Execute(filedNames, filedValues)


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
