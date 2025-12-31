Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Interface

Namespace PictureCaptionCrossReference
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

			' Add a paragraph to the section for the cross-reference
			Dim firstPara As Paragraph = section.AddParagraph()

			' Add another paragraph to the section
			Dim par1 As Paragraph = section.AddParagraph()
			par1.Format.AfterSpacing = 10

			' Append an image (picture) to the paragraph from the specified file path
			Dim pic1 As DocPicture = par1.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Spire.Doc.png"))
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim pic1 As DocPicture = par1.AppendPicture("..\..\..\..\..\..\Data\Spire.Doc.png")
			' =============================================================================
			pic1.Height = 120
			pic1.Width = 120

			' Set the caption numbering format to "Number" and add a caption below the picture
			Dim format As CaptionNumberingFormat = CaptionNumberingFormat.Number
			Dim captionParagraph As IParagraph = pic1.AddCaption("Figure", format, CaptionPosition.BelowItem)

			' Add another paragraph to the section
			Dim par2 As Paragraph = section.AddParagraph()

			' Append another image (picture) to the paragraph from the specified file path
			Dim pic2 As DocPicture = par2.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Word.png"))
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim pic2 As DocPicture = par2.AppendPicture("..\..\..\..\..\..\Data\Word.png")
			' =============================================================================
			pic2.Height = 120
			pic2.Width = 120

			' Add a caption below the second picture
			captionParagraph = pic2.AddCaption("Figure", format, CaptionPosition.BelowItem)

			' Add a bookmark at the specified location
			Dim bookmarkName As String = "Figure_2"
			Dim paragraph As Paragraph = section.AddParagraph()
			paragraph.AppendBookmarkStart(bookmarkName)
			paragraph.AppendBookmarkEnd(bookmarkName)

			' Navigate to the bookmark and replace its content with the caption paragraph
			Dim navigator As New BookmarksNavigator(document)
			navigator.MoveToBookmark(bookmarkName)
			Dim part As TextBodyPart = navigator.GetBookmarkContent()
			part.BodyItems.Clear()
			part.BodyItems.Add(captionParagraph)
			navigator.ReplaceBookmarkContent(part)

			' Create a cross-reference field for the bookmark
			Dim field As New Field(document)
			field.Type = FieldType.FieldRef
			field.Code = "REF Figure_2 \p \h"
			firstPara.ChildObjects.Add(field)
			Dim fieldSeparator As New FieldMark(document, FieldMarkType.FieldSeparator)
			firstPara.ChildObjects.Add(fieldSeparator)

			' Add the text "Figure 2" as the reference text
			Dim tr As New TextRange(document)
			tr.Text = "Figure 2"
			firstPara.ChildObjects.Add(tr)

			Dim fieldEnd As New FieldMark(document, FieldMarkType.FieldEnd)
			firstPara.ChildObjects.Add(fieldEnd)

			' Enable field updating in the document
			document.IsUpdateFields = True

			' Specify the output file name and format (Docx)
			Dim output As String = "PictureCaptionCrossReference.docx"
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
