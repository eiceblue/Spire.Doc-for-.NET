Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetTransparencyForTextbox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
			Dim doc As New Document()

			'Create a new section
			Dim section As Section = doc.AddSection()

			'Create a new paragraph
			Dim paragraph As Paragraph = section.AddParagraph()

			'Append TextBox
			Dim textbox1 As Spire.Doc.Fields.TextBox = paragraph.AppendTextBox(100, 50)

			'Set fill color
			textbox1.Format.FillColor = Color.Red

			'Set fill transparency
			textbox1.FillTransparency = 0.45

			'Save the Word file
			Dim output As String = "SetTransparencyForTextbox.docx"
			doc.SaveToFile(output, FileFormat.Docx2013)

			'Dispose the document
			doc.Dispose()
			
			'Launch the file 
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
