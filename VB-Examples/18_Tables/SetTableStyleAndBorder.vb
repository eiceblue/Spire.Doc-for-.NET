Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetTableStyleAndBorder
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
		   ' Create a new document object
		   Dim document As New Document()

		   ' Load a document from a file, specified by the file path
		   document.LoadFromFile("..\..\..\..\..\..\Data\TableSample.docx")

		   ' Get the first section of the document
		   Dim section As Section = document.Sections(0)

		   ' Get the first table in the section
		   Dim table As Table = TryCast(section.Tables(0), Table)

		   ' Apply a predefined table style to the table
		   table.ApplyStyle(DefaultTableStyle.ColorfulList)

		   ' Set the right border of the table
		   table.Format.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.Hairline
		   table.Format.Borders.Right.LineWidth = 1.0F
		   table.Format.Borders.Right.Color = Color.Red

		   ' Set the top border of the table
		   table.Format.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Hairline
		   table.Format.Borders.Top.LineWidth = 1.0F
		   table.Format.Borders.Top.Color = Color.Green

		   ' Set the left border of the table
		   table.Format.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.Hairline
		   table.Format.Borders.Left.LineWidth = 1.0F
		   table.Format.Borders.Left.Color = Color.Yellow

		   ' Set the bottom border of the table
		   table.Format.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.DotDash

		   ' Set the vertical borders of the table
		   table.Format.Borders.Vertical.BorderType = Spire.Doc.Documents.BorderStyle.Dot
		   table.Format.Borders.Vertical.Color = Color.Orange

		   ' Set the horizontal borders of the table to none
		   table.Format.Borders.Horizontal.BorderType = Spire.Doc.Documents.BorderStyle.None

		   ' Save the modified document to a file named "TableStyleAndBorder.docx", using Docx format
		   document.SaveToFile("TableStyleAndBorder.docx", FileFormat.Docx)

		   ' Dispose of the document object
		   document.Dispose() 
	   
			FileViewer("TableStyleAndBorder.docx")
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
