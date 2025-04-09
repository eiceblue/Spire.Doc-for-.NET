Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Pages
Imports Spire.Doc.Printing

Namespace PrintMultipageToOneSheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim inputFile As String = "..\..\..\..\..\..\Data\Template_Docx_4.docx"

			' Create a new instance of Document
			Dim doc As New Document()

			' Load the Word document from the specified input file
			doc.LoadFromFile(inputFile, FileFormat.Docx)

			' Create a new PrintDialog from System.Windows.Forms
			Dim printDialog As New PrintDialog()

			' Enable printing to a file
			printDialog.PrinterSettings.PrintToFile = True

			' Set the print file name based on the PagesPreSheet value
			printDialog.PrinterSettings.PrintFileName = String.Format("F:\TP\214\sample-new2.xps")

			' Assign the PrintDialog to the document's PrintDialog
			doc.PrintDialog = printDialog

			' Print the document with multiple pages condensed into one sheet
			doc.PrintMultipageToOneSheet(PagesPerSheet.FourPages, True)

			doc.Dispose()

			Me.Close()

		End Sub
	End Class
End Namespace
