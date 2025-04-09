Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace LockedAspectRatio
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim document As New Document()

			' Add a new section to the document
			Dim section As Section = document.AddSection()

			' Add a paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append a textbox to the paragraph and get a reference to it
			Dim textBox1 As Spire.Doc.Fields.TextBox = paragraph.AppendTextBox(240, 35)

			' Configure the horizontal alignment, line color, and line style of the textbox
			textBox1.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
			textBox1.Format.LineColor = Color.Black
			textBox1.Format.LineStyle = TextBoxLineStyle.Simple

			' Lock the aspect ratio of the textbox
			textBox1.AspectRatioLocked = True

			' Add a paragraph to the body of the textbox and get a reference to it
			Dim para As Paragraph = textBox1.Body.AddParagraph()

			' Add text to the paragraph
			Dim txtrg As TextRange = para.AppendText("Textbox 1 in the document")
			txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
			txtrg.CharacterFormat.FontSize = 14
			txtrg.CharacterFormat.TextColor = Color.Black
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			' Save the document to a file named "Sample.docx" in Docx format
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Dispose the document object to release resources
			document.Dispose()
			
			'Launch the  Word file.
			WordDocViewer("Sample.docx")
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
