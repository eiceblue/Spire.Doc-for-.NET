Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc

Namespace ToMarkdownWithOptions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document instance.
			Dim document As New Document()

			' Load the content of an existing .docx file.
			document.LoadFromFile("..\..\..\..\..\..\Data\toMarkdown.docx")

			' Configure the Markdown export options to embed images as Base64-encoded strings instead of saving them as separate files.
			' document.MarkdownExportOptions.ImagesAsBase64 = true;

			' Specify the local directory where extracted images from the document will be saved during Markdown conversion.
			document.MarkdownExportOptions.ImagesFolder = "D:\Markdown\Images"

			' Set an alias path for the images folder that will be used in the generated Markdown file.
			' document.MarkdownExportOptions.ImagesFolderAlias= "D:\\Markdown\\MyImages";

			' Set the list output mode to render bullet and numbered lists as plain text without Markdown list syntax.
			document.MarkdownExportOptions.ListOutputMode= MarkdownListOutputMode.PlainText

			' Enable preservation of underline formatting in the exported Markdown content.
			document.MarkdownExportOptions.SaveUnderlineFormatting = True

			' Configure hyperlink output to use Markdown reference-style links (e.g., [text][1] with definitions at the bottom).
			document.MarkdownExportOptions.LinkOutputMode = MarkdownLinkOutputMode.Reference

			' Set the output format for Office Math equations to MathML within the Markdown file.
			document.MarkdownExportOptions.OfficeMathOutputMode= MarkdownOfficeMathOutputMode.MathML

			' Specify that certain elements (e.g., tables) should be saved using HTML syntax within the Markdown output.
			document.MarkdownExportOptions.SaveAsHtml = MarkdownSaveAsHtml.Tables

			' Save the converted document to a Markdown file.
			document.SaveToFile("toMarkdown_out.md", FileFormat.Markdown)

			' Close the document to release associated resources.
			document.Close()

			' Explicitly dispose of the document object to free memory and system resources.
			document.Dispose()
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
