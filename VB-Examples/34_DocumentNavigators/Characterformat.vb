Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace Characterformat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Load the document by creating a new Document instance.
			Dim doc As New Document()

			' Create a DocumentNavigator object to facilitate easy content insertion and formatting.
			Dim navigator As New DocumentNavigator(doc)

			' Write plain text into the document (without special formatting yet).
			navigator.Write("Write plain text into the document (without special formatting yet).")

			' Insert a line break after the previous text.
			navigator.InsertBreak(BreakType.LineBreak)

			' Enable underline formatting for subsequent text.
			navigator.CharacterFormat.UnderlineColor = Color.OrangeRed
			navigator.CharacterFormat.UnderlineStyle = UnderlineStyle.Single

			' Set bold formatting for subsequent text.
			navigator.CharacterFormat.Bold = True

			' Enable shadow effect for subsequent text.
			navigator.CharacterFormat.IsShadow = True

			' Set the text color to blue for subsequent text.
			navigator.CharacterFormat.TextColor = Color.Blue

			' Write formatted text using the current character formatting settings.
			navigator.Write("Write formatted text using the current character formatting settings.")

			' Insert another line break.
			navigator.InsertBreak(BreakType.LineBreak)

			' Save the current character formatting settings onto an internal stack for later reuse.
			navigator.PushCharacterFormat()

			' Clear all character formatting to default (e.g., no bold, no color, etc.).
			navigator.CharacterFormat.ClearFormatting()

			' Write text with cleared (default) formatting.
			navigator.Write("Write text with cleared (default) formatting")

			' Insert another line break.
			navigator.InsertBreak(BreakType.LineBreak)

			' Restore the previously saved character formatting from the stack.
			navigator.PopCharacterFormat()

			' Write text using the restored formatting.
			navigator.Write("Write text using the restored formatting.")

			' Save the document to a file in DOCX format.
			doc.SaveToFile("Characterformat.docx", FileFormat.Docx)

			' Close the document to release internal resources.
			doc.Close()

			' Explicitly dispose of the document object to free memory immediately.
			doc.Dispose()

			'Launch the MS Word file.
			WordDocViewer("Characterformat.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
