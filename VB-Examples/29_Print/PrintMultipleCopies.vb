Imports Spire.Doc

Namespace PrintMultipleCopies
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim document As New Document()

			' Load the Word document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			' Set the printer name to "Microsoft Print to PDF" for printing
			document.PrintDocument.PrinterSettings.PrinterName = "Microsoft Print to PDF"

			' Set the number of copies to be printed to 4
			document.PrintDocument.PrinterSettings.Copies = 4

			' Print the document
			document.PrintDocument.Print()

			' Dispose of the document object when finished using it
			document.Dispose()

		End Sub

	End Class
End Namespace
