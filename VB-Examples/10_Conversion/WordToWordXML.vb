Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace WordToWordXML
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document.
			Dim document As New Document()

			'Load the file from disk.
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			Dim result1 As String = "Result-WordToWordML.xml"
			Dim result2 As String = "Result-WordToWordXML.xml"

			'For word 2003:
			document.SaveToFile(result1, FileFormat.WordML)

			'For word 2007:
			document.SaveToFile(result2, FileFormat.WordXml)

			'Dispose the document.
			document.Dispose()

			'Launch the files.
			WordDocViewer(result1)
			WordDocViewer(result2)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
