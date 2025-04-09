Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text

Namespace AddImageToEachPage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Load a Word document
			Dim document As New Document("..\..\..\..\..\..\Data\SampleB_2.docx")

			'Specified the image path
			Dim imgPath As String = "..\..\..\..\..\..\Data\Spire.Doc.png"

			'Add a picture in footer
			Dim picture As DocPicture = document.Sections(0).HeadersFooters.Footer.AddParagraph().AppendPicture(Image.FromFile(imgPath))

			'Set the picture's postion and style
			picture.VerticalOrigin = VerticalOrigin.Page
			picture.HorizontalOrigin = HorizontalOrigin.Page
			picture.VerticalAlignment = ShapeVerticalAlignment.Bottom
			picture.TextWrappingStyle = TextWrappingStyle.None

			'Add a textbox in footer 
			Dim textbox As Spire.Doc.Fields.TextBox = document.Sections(0).HeadersFooters.Footer.AddParagraph().AppendTextBox(150, 20)

			'Set the textbox's postion and style
			textbox.VerticalOrigin = VerticalOrigin.Page
			textbox.HorizontalOrigin = HorizontalOrigin.Page
			textbox.HorizontalPosition = 300
			textbox.VerticalPosition = 700
			textbox.Body.AddParagraph().AppendText("Welcome to E-iceblue")

			'Save to file
			document.SaveToFile("result.docx", FileFormat.Docx)

			'Dispose the document
			document.Dispose()

			'Launch result file
			WordDocViewer("result.docx")

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
