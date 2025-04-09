Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetHyperlinkFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path for the document
			Dim input As String = "..\..\..\..\..\..\Data\BlankTemplate.docx"

			' Create a new Document object
			Dim doc As New Document()

			' Load the document from the specified file path
			doc.LoadFromFile(input)

			' Get the first section of the document
			Dim section As Section = doc.Sections(0)

			' Add a paragraph to the section and append regular text
			Dim para1 As Paragraph = section.AddParagraph()
			para1.AppendText("Regular Link: ")

			' Append a hyperlink to the paragraph with the specified URL and display text
			Dim txtRange1 As TextRange = para1.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)
			txtRange1.CharacterFormat.FontName = "Times New Roman"
			txtRange1.CharacterFormat.FontSize = 12

			' Add a blank paragraph as separation
			Dim blankPara1 As Paragraph = section.AddParagraph()

			' Add another paragraph to the section and append text
			Dim para2 As Paragraph = section.AddParagraph()
			para2.AppendText("Change Color: ")

			' Append a hyperlink to the paragraph with the specified URL and display text
			Dim txtRange2 As TextRange = para2.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)
			txtRange2.CharacterFormat.FontName = "Times New Roman"
			txtRange2.CharacterFormat.FontSize = 12
			txtRange2.CharacterFormat.TextColor = Color.Red

			' Add a blank paragraph as separation
			Dim blankPara2 As Paragraph = section.AddParagraph()

			' Add another paragraph to the section and append text
			Dim para3 As Paragraph = section.AddParagraph()
			para3.AppendText("Remove Underline: ")

			' Append a hyperlink to the paragraph with the specified URL and display text
			Dim txtRange3 As TextRange = para3.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)
			txtRange3.CharacterFormat.FontName = "Times New Roman"
			txtRange3.CharacterFormat.FontSize = 12
			txtRange3.CharacterFormat.UnderlineStyle = UnderlineStyle.None

			' Specify the output file path for the modified document
			Dim output As String = "HyperlinkFormat.docx"

			' Save the modified document to the output file path in DOCX format
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose the document object to free up resources
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
