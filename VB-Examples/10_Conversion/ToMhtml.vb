Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ToMhtml
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
			Dim document As New Document()

			'Load the file from disk.
			document.LoadFromFile("..\..\..\..\..\..\Data\ToMhtml.docx")

			'Save to RTF file.
			document.SaveToFile("ToMhtml-out.mhtml", FileFormat.Mhtml)

			'Dispose the document
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("ToMhtml-out.mhtml")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
