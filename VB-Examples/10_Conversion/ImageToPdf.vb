Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO
Imports Spire.Doc.Fields
Namespace ImageToPdf
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\Image.png"

			'Create a new document
			Dim doc As New Document()

			'Create a new section
			Dim section As Section = doc.AddSection()

			'Create a new paragraph
			Dim paragraph As Paragraph = section.AddParagraph()

			'Add a picture for paragraph
			Dim picture As DocPicture = paragraph.AppendPicture(input)

			'Set A4 page size
			section.PageSetup.PageSize = PageSize.A4

			'Set the page margins
			section.PageSetup.Margins.Top = 10f
			section.PageSetup.Margins.Left = 25f

			'Save to file
			Dim result As String = "ImageToPdf.pdf"
			doc.SaveToFile(result,FileFormat.PDF)

			'Dispose the Document
			doc.Dispose()

			Viewer(result)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
