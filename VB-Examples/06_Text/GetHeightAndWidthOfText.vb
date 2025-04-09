Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.IO

Namespace GetHeightAndWidthOfText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object.
			Dim document As New Document()

			' Load a Word document from the specified file path.
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_2.docx")

			' Specify a text to search for in the document.
			Dim text As String = "Your Office Development Master"

			' Find the first occurrence of the specified text in the document and obtain a TextSelection object.
			Dim selection As TextSelection = document.FindString(text, True, True)

			' Get the font used for the selected text range.
			Dim font As Font = selection.GetAsOneRange().CharacterFormat.Font

			' Create a fake image with a size of 1x1 pixel.
			Dim fakeImage As Image = New Bitmap(1, 1)

			' Create a Graphics object from the fake image.
			Dim graphics As Graphics = Graphics.FromImage(fakeImage)

			' Measure the size of the specified text using the selected font.
			Dim size As SizeF = graphics.MeasureString(text, font)

			' Create a StringBuilder object to store the content.
			Dim content As New StringBuilder()

			' Append the text height and width to the content.
			content.AppendLine("text height: " & size.Height)
			content.AppendLine("text width: " & size.Width)

			' Specify the output file name.
			Dim result As String = "Result-GetHeightAndWidthOfText.txt"

			' Write the content to the output file.
			File.WriteAllText(result, content.ToString())

			' Dispose of the Document object to release resources.
			document.Dispose()

			'Launch the file.
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
