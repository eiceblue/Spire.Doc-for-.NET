Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetIndentByCharacter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new Document object
			Dim document As New Document()

			' Add a section to the document
			Dim sec As Section = document.AddSection()

			' Add a paragraph for the title
			Dim para As Paragraph = sec.AddParagraph()
			para.AppendText("Paragraph Formatting")
			para.ApplyStyle(BuiltinStyle.Title)

			' Add a paragraph with indent settings
			para = sec.AddParagraph()
			para.AppendText("This paragraph is indent as follows: Indent 2 characters on the left and 5 characters on the right.")
			para.Format.LeftIndentChars= 2f
			para.Format.RightIndentChars = 5f

			' Specify the output file name for the modified document
			Dim output As String = "SetIndentByCharacter_output.docx"

		   ' Save the modified document to the specified file format
		   document.SaveToFile(output, FileFormat.Docx)

		   ' Dispose the Document object to free resources
		  document.Dispose()

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
