Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemoveFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\HeaderAndFooter.docx"

			'Create a Word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first section
			Dim section As Section = doc.Sections(0)

			'Traverse the word document and clear all footers in different type
			For Each para As Paragraph In section.Paragraphs

				'Loop through the child objects of paragraph
				For Each obj As DocumentObject In para.ChildObjects
					'Clear footer in the first page
					Dim footer As HeaderFooter
					footer = section.HeadersFooters(HeaderFooterType.FooterFirstPage)
					If footer IsNot Nothing Then
						footer.ChildObjects.Clear()
					End If
					'Clear footer in the odd page
					footer = section.HeadersFooters(HeaderFooterType.FooterOdd)
					If footer IsNot Nothing Then
						footer.ChildObjects.Clear()
					End If
					'Clear footer in the even page
					footer = section.HeadersFooters(HeaderFooterType.FooterEven)
					If footer IsNot Nothing Then
						footer.ChildObjects.Clear()
					End If
				Next obj
			Next para

			'Save the document
			Dim output As String = "RemoveFooter.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
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
