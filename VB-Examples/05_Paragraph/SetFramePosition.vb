Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.IO
Imports System.Text

Namespace SetFramePosition
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\TextInFrame.docx")

			' Retrieve the first paragraph in the first section of the document
			Dim paragraph As Paragraph = document.Sections(0).Paragraphs(0)

			' Check if the paragraph has a frame formatting
			If paragraph.Frame.IsFrame Then
				' Set the horizontal position of the frame to 150.0F
				paragraph.Frame.SetHorizontalPosition(150.0F)
				
				' Set the vertical position of the frame to 150.0F
				paragraph.Frame.SetVerticalPosition(150.0F)
			End If

			' Specify the output file name for the modified document
			Dim result As String = "SetFramePosition_result.docx"

			' Save the document with the modified frame position to the specified file format (Docx2013)
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object to release resources
			document.Dispose()

			'Launch the file
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
