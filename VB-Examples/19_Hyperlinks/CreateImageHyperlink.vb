Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace CreateImageHyperlink
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
		   ' Specify the input file path for the template document
		   Dim input As String = "..\..\..\..\..\..\Data\BlankTemplate.docx"

		   ' Create a new Document object
		   Dim doc As New Document()

		   ' Load the template document from the specified file path
		   doc.LoadFromFile(input)

		   ' Get the first section of the document
		   Dim section As Section = doc.Sections(0)

		   ' Add a new paragraph in the section
		   Dim paragraph As Paragraph = section.AddParagraph()

		   ' Load an image from the specified file path
		   Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\Spire.Doc.png")

		   ' Create a new DocPicture object with the loaded image
		   Dim picture As New DocPicture(doc)

			' Load the image into the DocPicture object
			picture.LoadImage(image)
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim picture As New DocPicture(doc)
			'picture.LoadImage("..\..\..\..\..\..\Data\Spire.Doc.png")
			' =============================================================================


			' Append a hyperlink to the paragraph with the specified URL and the picture as the display element
			paragraph.AppendHyperlink("https://www.e-iceblue.com/Introduce/word-for-net-introduce.html", picture, HyperlinkType.WebLink)

		   ' Specify the output file path for the generated document
		   Dim output As String = "CreateImageHyperlink.docx"

		   ' Save the document to the output file path in DOCX format
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
