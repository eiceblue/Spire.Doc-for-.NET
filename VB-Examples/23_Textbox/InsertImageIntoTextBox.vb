Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace InsertImageIntoTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim doc As New Document()

			' Add a section to the document
			Dim section As Section = doc.AddSection()

			' Add a paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append a text box to the paragraph with specified dimensions
			Dim tb As Spire.Doc.Fields.TextBox = paragraph.AppendTextBox(220, 220)

			' Set the horizontal and vertical positioning of the text box
			tb.Format.HorizontalOrigin = HorizontalOrigin.Page
			tb.Format.HorizontalPosition = 50
			tb.Format.VerticalOrigin = VerticalOrigin.Page
			tb.Format.VerticalPosition = 50

			' Set the background fill effect of the text box to a picture
			tb.Format.FillEfects.Type = BackgroundType.Picture

			' Set the picture for the background fill effect from a file
			tb.Format.FillEfects.Picture = Image.FromFile("..\..\..\..\..\..\Data\Spire.Doc.png")

			' Specify the output file name
			Dim output As String = "InsertImageIntoTextBox.docx"

			' Save the document to a file in DOCX format
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose the Document object to free up resources
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
