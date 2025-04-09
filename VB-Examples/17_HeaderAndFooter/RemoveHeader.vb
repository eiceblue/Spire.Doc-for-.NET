Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemoveHeader
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

			'Get the first section of the document
			Dim section As Section = doc.Sections(0)

			'Traverse the word document and clear all headers in different type
			For Each para As Paragraph In section.Paragraphs
				For Each obj As DocumentObject In para.ChildObjects
					'Clear footer in the first page
					Dim header As HeaderFooter
					header = section.HeadersFooters(HeaderFooterType.HeaderFirstPage)
					If header IsNot Nothing Then
						header.ChildObjects.Clear()
					End If
					'Clear footer in the odd page
					header = section.HeadersFooters(HeaderFooterType.HeaderOdd)
					If header IsNot Nothing Then
						header.ChildObjects.Clear()
					End If
					'Clear footer in the even page
					header = section.HeadersFooters(HeaderFooterType.HeaderEven)
					If header IsNot Nothing Then
						header.ChildObjects.Clear()
					End If
				Next obj
			Next para

			'Save the document
			Dim output As String = "RemoveHeader.docx"
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
