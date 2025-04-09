Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace CopyHeaderAndFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\HeaderAndFooter.docx"

			'Create a word document
			Dim doc1 As New Document()

			'Load the file from disk
			doc1.LoadFromFile(input)

			'Get the header section from the source document
			Dim header As HeaderFooter = doc1.Sections(0).HeadersFooters.Header

			input = "..\..\..\..\..\..\Data\Template.docx"

			'Create another Word document
			Dim doc2 As New Document()

			'Load the destination file
			doc2.LoadFromFile(input)

			'Copy each object in the header of source file to destination file
			For Each section As Section In doc2.Sections

				'Loop through the child objects of heder
				For Each obj As DocumentObject In header.ChildObjects

					'Copy each object in the header of source file to destination file
					section.HeadersFooters.Header.ChildObjects.Add(obj.Clone())
				Next obj
			Next section

			'Save the document
			Dim output As String = "CopyHeaderAndFooter.docx"
			doc2.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
			doc1.Dispose()
			doc2.Dispose()
			
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
