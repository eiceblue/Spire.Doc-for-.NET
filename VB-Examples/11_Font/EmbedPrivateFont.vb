Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace EmbedPrivateFont
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\BlankTemplate.docx"

			'Create a Word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first section
			Dim section As Section = doc.Sections(0)

			'Get the first paragraph
			Dim p As Paragraph = section.AddParagraph()

			'Append text to the paragraph
			Dim range As TextRange = p.AppendText("Spire.Doc for .NET is a professional Word.NET library specifically designed for developers to create, read, write, convert and print Word document files from any.NET platform with fast and high quality performance.")

			'Set the font name and font size
			range.CharacterFormat.FontName = "PT Serif Caption"
			range.CharacterFormat.FontSize = 20

			'Allow embedding font in document
			doc.EmbedFontsInFile = True

			'Embed private font from font file into the document
			doc.PrivateFontList.Add(New PrivateFontPath("PT Serif Caption", "..\..\..\..\..\..\Data\PT Serif Caption.ttf"))

			'Save the document
			Dim output As String = "EmbedPrivateFont.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			'Dispose the document.
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
