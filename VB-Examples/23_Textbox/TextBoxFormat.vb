Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace TextBoxFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim doc As New Document()

			' Add a new section to the document
			Dim sec As Section = doc.AddSection()

			' Add a textbox to the first paragraph in the section and get a reference to it
			Dim TB As Spire.Doc.Fields.TextBox = doc.Sections(0).AddParagraph().AppendTextBox(310, 90)

			' Add a paragraph to the body of the textbox and get a reference to it
			Dim para As Paragraph = TB.Body.AddParagraph()

			' Add text to the paragraph
			Dim TR As TextRange = para.AppendText("Using Spire.Doc, developers will find " & "a simple and effective method to endow their applications with rich MS Word features. ")

			' Set the font properties for the text
			TR.CharacterFormat.FontName = "Cambria"
			TR.CharacterFormat.FontSize = 13

			' Configure the position of the textbox
			TB.Format.HorizontalOrigin = HorizontalOrigin.Page
			TB.Format.HorizontalPosition = 120
			TB.Format.VerticalOrigin = VerticalOrigin.Page
			TB.Format.VerticalPosition = 100

			' Configure the line style and color of the textbox
			TB.Format.LineStyle = TextBoxLineStyle.Double
			TB.Format.LineColor = Color.CornflowerBlue
			TB.Format.LineDashing = LineDashing.Solid
			TB.Format.LineWidth = 5

			' Configure the internal margins of the textbox
			TB.Format.InternalMargin.Top = 15
			TB.Format.InternalMargin.Bottom = 10
			TB.Format.InternalMargin.Left = 12
			TB.Format.InternalMargin.Right = 10

			' Specify the output file path
			Dim output As String = "TextBoxFormat.docx"

			' Save the modified document to the output file with the specified file format (Docx)
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose the document object to release resources
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
