Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RetrieveStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Load a template document 
			Dim doc As New Document("..\..\..\..\..\..\Data\Styles.docx")

			'Traverse all paragraphs in the document and get their style names through StyleName property
			Dim styleName As String = Nothing
			For Each section As Section In doc.Sections
				For Each paragraph As Paragraph In section.Paragraphs
					styleName &= paragraph.StyleName & vbCrLf
				Next paragraph
			Next section

			'Save the style name as a txt file
			Dim output As String = "RetrieveStyle.txt"
			File.WriteAllText(output, styleName.ToString())

			'Dispose of the document object
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
