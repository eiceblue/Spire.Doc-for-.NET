Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace CreateBookmark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Open a blank word document as template
			Dim document As New Document("..\..\..\..\..\..\..\Data\Blank.doc")

			CreateBookmark(document.Sections(0))

			'Save doc file.
			document.SaveToFile("Sample.doc",FileFormat.Doc)

			'Launching the MS Word file.
			WordDocViewer("Sample.doc")


		End Sub

		Private Sub CreateBookmark(ByVal section As Section)
            Dim paragraph As Paragraph
            If (section.Paragraphs.Count > 0) Then
                paragraph = section.Paragraphs(0)
            Else
                paragraph = section.AddParagraph()
            End If
            paragraph.AppendText("The sample demonstrates how to using CreateBookmark.")
            paragraph.ApplyStyle(BuiltinStyle.Heading2)

            section.AddParagraph()
            paragraph = section.AddParagraph()
            paragraph.AppendText("Simple CreateBookmark.")
            paragraph.ApplyStyle(BuiltinStyle.Heading4)

            ' Writing simple CreateBookmarks
            section.AddParagraph()
            paragraph = section.AddParagraph()
            paragraph.AppendBookmarkStart("SimpleCreateBookmark")
            paragraph.AppendText("This is a simple book mark.")
            paragraph.AppendBookmarkEnd("SimpleCreateBookmark")

            section.AddParagraph()
            paragraph = section.AddParagraph()
            paragraph.AppendText("Nested CreateBookmark.")
            paragraph.ApplyStyle(BuiltinStyle.Heading4)

            ' Writing nested CreateBookmarks
            section.AddParagraph()
            paragraph = section.AddParagraph()
            paragraph.AppendBookmarkStart("Root")
            paragraph.AppendText(" Root data ")
            paragraph.AppendBookmarkStart("NestedLevel1")
            paragraph.AppendText(" Nested Level1 ")
            paragraph.AppendBookmarkStart("NestedLevel2")
            paragraph.AppendText(" Nested Level2 ")
            paragraph.AppendBookmarkEnd("NestedLevel2")
            paragraph.AppendText(" Data Level1 ")
            paragraph.AppendBookmarkEnd("NestedLevel1")
            paragraph.AppendText(" Data Root ")
            paragraph.AppendBookmarkEnd("Root")

        End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
