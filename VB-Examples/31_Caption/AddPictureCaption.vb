Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddPictureCaption
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
			Dim par1 As Paragraph = section.AddParagraph()
			par1.Format.AfterSpacing = 10

			' Append an image (picture) to the paragraph from the specified file path
			Dim pic1 As DocPicture = par1.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Spire.Doc.png"))
			pic1.Height = 100
			pic1.Width = 120

			' Set the caption numbering format to "Number" and add a caption below the picture
			Dim format As CaptionNumberingFormat = CaptionNumberingFormat.Number
			pic1.AddCaption("Figure", format, CaptionPosition.BelowItem)

			' Add another paragraph to the section
			Dim par2 As Paragraph = section.AddParagraph()

			' Append another image (picture) to the paragraph from the specified file path
			Dim pic2 As DocPicture = par2.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Word.png"))
			pic2.Height = 100
			pic2.Width = 120

			' Add a caption below the second picture
			pic2.AddCaption("Figure", format, CaptionPosition.BelowItem)

			' Enable field updating in the document
			document.IsUpdateFields = True

			' Specify the output file name and format (Docx)
			Dim output As String = "AddPictureCaption_result.docx"
			document.SaveToFile(output, FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launching the file
			WordDocViewer(output)

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
