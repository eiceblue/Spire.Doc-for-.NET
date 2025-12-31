Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertOLE
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim doc As New Document()

			' Add a section to the document
			Dim sec As Section = doc.AddSection()

			' Add a paragraph to the section
			Dim par As Paragraph = sec.AddParagraph()

			' Create a DocPicture object and load an image from file
			Dim picture As New DocPicture(doc)
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\excel.png")
			picture.LoadImage(image)
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'picture.LoadImage("..\..\..\..\..\..\Data\excel.png")
			' =============================================================================

			' Append an OLE object to the paragraph with the specified file, picture, and object type (Excel worksheet)
			Dim obj As DocOleObject = par.AppendOleObject("..\..\..\..\..\..\Data\example.xlsx", picture, OleObjectType.ExcelWorksheet)

			' Save the document to a file in Docx2013 format
			doc.SaveToFile("InsertOLE.docx", FileFormat.Docx2013)

			' Dispose the document object
			doc.Dispose()

			FileViewer("InsertOLE.docx")
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
