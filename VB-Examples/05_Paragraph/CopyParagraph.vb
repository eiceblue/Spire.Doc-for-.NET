Imports System.ComponentModel
Imports System.Net.Mime.MediaTypeNames
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Namespace CopyParagraph
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object called document1
			Dim document1 As New Document()

			'Load a Word document from the specified file path
			document1.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_5.docx")

			'Create a new Document object called document2
			Dim document2 As New Document()

			'Get the first section, first paragraph, and second paragraph from document1 and assign them to variables
			Dim s As Section = document1.Sections(0)
			Dim p1 As Paragraph = s.Paragraphs(0)
			Dim p2 As Paragraph = s.Paragraphs(1)

			'Add a new section to document2 and assign it to s2
			Dim s2 As Section = document2.AddSection()

			'Clone the first paragraph and add it to the paragraphs collection of s2
			Dim NewPara1 As Paragraph = CType(p1.Clone(), Paragraph)
			s2.Paragraphs.Add(NewPara1)

			'Clone the second paragraph and add it to the paragraphs collection of s2
			Dim NewPara2 As Paragraph = CType(p2.Clone(), Paragraph)
			s2.Paragraphs.Add(NewPara2)

			'Create a new PictureWatermark object and assign it to WM
			Dim WM As New PictureWatermark()

			'Load an image from the specified file path and assign it to the Picture property of WM
			WM.Picture = Image.FromFile("..\..\..\..\..\..\Data\Logo.jpg")
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'WM.Picture("..\..\..\..\..\..\Data\Logo.jpg")
			' =============================================================================

			'Set the watermark of document2 to WM
			document2.Watermark = WM

			'Specify the output file name
			Dim result As String = "Result-CopyWordParagraph.docx"

			'Save document2 as a Word document with the specified file format
			document2.SaveToFile(result, FileFormat.Docx2013)

			'Dispose of document1 object to release resources
			document1.Dispose()

			'Dispose of document2 object to release resources
			document2.Dispose()

			'Launch the MS Word file.
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
