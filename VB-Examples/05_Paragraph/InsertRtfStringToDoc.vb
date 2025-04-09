Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace InsertRtfStringToDoc
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Add a Section to the document
			Dim section As Section = document.AddSection()

			' Add a Paragraph to the Section
			Dim para As Paragraph = section.AddParagraph()

			' Define an RTF string
			Dim rtfString As String = "{\rtf1\ansi\deff0 {\fonttbl {\f0 hakuyoxingshu7000;}}\f0\fs28 Hello, World}"

			' Append the RTF string to the Paragraph
			para.AppendRTF(rtfString)

			' Specify the filename for the output result
			Dim result As String = "Result-InsertRtfStringToWord.docx"

			' Save the document to the specified file in Docx format
			document.SaveToFile(result, FileFormat.Docx)

			' Dispose the Document object to release resources
			document.Dispose()

			'Launch the MS Word file.
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
