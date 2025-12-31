Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ImageHeaderAndFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\Template.docx"

			'Create a word document
			Dim doc As New Document()

			'Load the document from disk
			doc.LoadFromFile(input)

			'Get the header of the first page
			Dim header As HeaderFooter = doc.Sections(0).HeadersFooters.Header

			'Add a paragraph for the header
			Dim paragraph As Paragraph = header.AddParagraph()

			'Set the format of the paragraph
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right

			'Append a picture in the paragraph
			Dim headerimage As DocPicture = paragraph.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\E-iceblue.png"))
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim headerimage As DocPicture = paragraph.AppendPicture("..\..\..\..\..\..\Data\E-iceblue.png")
			' =============================================================================
			headerimage.VerticalAlignment = ShapeVerticalAlignment.Bottom

			'Get the footer of the first section
			Dim footer As HeaderFooter = doc.Sections(0).HeadersFooters.Footer

			'Add a paragraph for the footer
			Dim paragraph2 As Paragraph = footer.AddParagraph()

			'Set the format of the paragraph
			paragraph2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left

			'Append a picture in the paragraph
			Dim footerimage As DocPicture = paragraph2.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\logo.png"))
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim footerimage As DocPicture = paragraph2.AppendPicture("..\..\..\..\..\..\Data\logo.png")
			' =============================================================================

			'Append text in the paragraph and set its character format
			Dim TR As TextRange = paragraph2.AppendText("Copyright © 2013 e-iceblue. All Rights Reserved.")
			TR.CharacterFormat.FontName = "Arial"
			TR.CharacterFormat.FontSize = 10
			TR.CharacterFormat.TextColor = Color.Black

			'Save and launch document
			Dim output As String = "ImageHeaderAndFooter.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
			doc.Dispose()
			
			Viewer(output)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
