Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertingTextbox
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

			' Call the method to insert a textbox into the section
			InsertTextbox(section)

			' Save the document to a file named "Sample.docx" in Docx format
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Dispose the document object to release resources
			document.Dispose()

			'Launching the Word file.
			WordDocViewer("Sample.docx")


		End Sub


		Private Shared Sub InsertTextbox(ByVal section As Section)
			' Create a paragraph in the specified section
			Dim paragraph As Paragraph = If(section.Paragraphs.Count > 0, section.Paragraphs(0), section.AddParagraph())

			' Add three paragraphs to create space between textboxes
			paragraph = section.AddParagraph()
			paragraph = section.AddParagraph()
			paragraph = section.AddParagraph()

			' Create and customize textbox 1
			Dim textBox1 As Spire.Doc.Fields.TextBox = paragraph.AppendTextBox(240, 35)
			textBox1.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
			textBox1.Format.LineColor = Color.Gray
			textBox1.Format.LineStyle = TextBoxLineStyle.Simple
			textBox1.Format.FillColor = Color.DarkSeaGreen
			Dim para As Paragraph = textBox1.Body.AddParagraph()
			Dim txtrg As TextRange = para.AppendText("Textbox 1 in the document")
			txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
			txtrg.CharacterFormat.FontSize = 14
			txtrg.CharacterFormat.TextColor = Color.White
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			' Add four paragraphs to create space between textboxes
			paragraph = section.AddParagraph()
			paragraph = section.AddParagraph()
			paragraph = section.AddParagraph()
			paragraph = section.AddParagraph()

			' Create and customize textbox 2
			Dim textBox2 As Spire.Doc.Fields.TextBox = paragraph.AppendTextBox(240, 35)
			textBox2.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
			textBox2.Format.LineColor = Color.Tomato
			textBox2.Format.LineStyle = TextBoxLineStyle.ThinThick
			textBox2.Format.FillColor = Color.Blue
			textBox2.Format.LineDashing = LineDashing.Dot
			para = textBox2.Body.AddParagraph()
			txtrg = para.AppendText("Textbox 2 in the document")
			txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
			txtrg.CharacterFormat.FontSize = 14
			txtrg.CharacterFormat.TextColor = Color.Pink
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			' Add four paragraphs to create space between textboxes
			paragraph = section.AddParagraph()
			paragraph = section.AddParagraph()
			paragraph = section.AddParagraph()
			paragraph = section.AddParagraph()

			' Create and customize textbox 3
			Dim textBox3 As Spire.Doc.Fields.TextBox = paragraph.AppendTextBox(240, 35)
			textBox3.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
			textBox3.Format.LineColor = Color.Violet
			textBox3.Format.LineStyle = TextBoxLineStyle.Triple
			textBox3.Format.FillColor = Color.Pink
			textBox3.Format.LineDashing = LineDashing.DashDotDot
			para = textBox3.Body.AddParagraph()
			txtrg = para.AppendText("Textbox 3 in the document")
			txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
			txtrg.CharacterFormat.FontSize = 14
			txtrg.CharacterFormat.TextColor = Color.Tomato
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
