Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO

Namespace PrintDocViaXpsPrint
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new MemoryStream for storing the document as XPS
			Using ms As New MemoryStream()
				' Instantiate a new Document object
				Using document As New Document()
					' Load the Word document from the specified template file
					document.LoadFromFile("..\..\..\..\..\..\Data\Template.docx")

					' Save the document to the MemoryStream as XPS format
					document.SaveToStream(ms, FileFormat.XPS)
				End Using

				' Reset the position of the MemoryStream to the beginning
				ms.Position = 0

				' Specify the printer name to be used for printing
				Dim printerName As String = "HP LaserJet P1007"

				' Print the XPS document using the specified printer and job name
				XpsPrint.XpsPrintHelper.Print(ms, printerName, "My printing job", True)
			End Using
		End Sub
	End Class
End Namespace
