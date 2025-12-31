Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace Hyperlink
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Add a section to the document
			Dim section As Section = document.AddSection()

			' Call the InsertHyperlink method to insert hyperlinks in the section
			InsertHyperlink(section)

			' Save the document to a file named "Sample.docx" in DOCX format
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Dispose the document object to free up resources
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("Sample.docx")


		End Sub

		' Define the InsertHyperlink method
		Private Shared Sub InsertHyperlink(ByVal section As Section)
			' Add a paragraph to the section, or get the first paragraph if it exists
			Dim paragraph As Paragraph = If(section.Paragraphs.Count > 0, section.Paragraphs(0), section.AddParagraph())

			' Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Spire.Doc for .NET " & vbCrLf & " e-iceblue company Ltd. 2002-2010 All rights reserverd")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			' Add a new paragraph to the section
			paragraph = section.AddParagraph()

			' Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Home page")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			' Add a hyperlink to the paragraph with the specified URL and display text
			paragraph = section.AddParagraph()
			paragraph.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)

			' Add a new paragraph to the section
			paragraph = section.AddParagraph()

			' Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Contact US")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			' Add a hyperlink to the paragraph with the specified email address and display text
			paragraph = section.AddParagraph()
			paragraph.AppendHyperlink("mailto:support@e-iceblue.com", "support@e-iceblue.com", HyperlinkType.EMailLink)

			' Add a new paragraph to the section
			paragraph = section.AddParagraph()

			' Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Forum")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			' Add a hyperlink to the paragraph with the specified URL and display text
			paragraph = section.AddParagraph()
			paragraph.AppendHyperlink("www.e-iceblue.com/forum/", "www.e-iceblue.com/forum/", HyperlinkType.WebLink)

			' Add a new paragraph to the section
			paragraph = section.AddParagraph()

			' Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Download Link")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			' Add a hyperlink to the paragraph with the specified URL and display text
			paragraph = section.AddParagraph()
			paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-net-now.html", "www.e-iceblue.com/Download/download-word-for-net-now.html", HyperlinkType.WebLink)

			' Add a new paragraph to the section
			paragraph = section.AddParagraph()

			' Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Insert Link On Image")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			' Add an image to the paragraph and append a hyperlink to it with the specified URL and link type
			paragraph = section.AddParagraph()
			Dim picture As DocPicture = paragraph.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Spire.Doc.png"))
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim picture As DocPicture = paragraph.AppendPicture("..\..\..\..\..\..\Data\Spire.Doc.png")
			' =============================================================================

			paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-net-now.html", picture, HyperlinkType.WebLink)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
