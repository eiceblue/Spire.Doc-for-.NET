Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.Drawing.Printing

Namespace SetMarginAndDuplex
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim input As String = "..\..\..\..\..\..\Data\Sample.docx"

			' Create a new instance of Document
			Dim doc As New Document()

			' Load the Word document from the specified input file
			doc.LoadFromFile(input)

			' Get the PrintDocument associated with the document
			Dim printDoc As PrintDocument = doc.PrintDocument

			' Set the OriginAtMargins property to true to align the printable area with the margins
			printDoc.OriginAtMargins = True

			' Set the Margins property of the DefaultPageSettings to zero to remove any margins
			printDoc.DefaultPageSettings.Margins = New System.Drawing.Printing.Margins(0, 0, 0, 0)

			' Set the Duplex property of PrinterSettings to Vertical for double-sided printing
			printDoc.PrinterSettings.Duplex = Duplex.Vertical

			' Print the document
			printDoc.Print()

			' Dispose of the document object when finished using it
			doc.Dispose()
		End Sub
	End Class
End Namespace
