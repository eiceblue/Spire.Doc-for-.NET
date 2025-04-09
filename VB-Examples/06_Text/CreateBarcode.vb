Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace CreateBarcode
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim doc As New Document()

			' Add a section to the document and get a reference to the newly added paragraph
			Dim p As Paragraph = doc.AddSection().AddParagraph()

			' Create a TextRange object and set its text to "H63TWX11072"
			Dim txtRang As TextRange = p.AppendText("H63TWX11072")

			' Set the font name of the text to "C39HrP60DlTt"
			txtRang.CharacterFormat.FontName = "C39HrP60DlTt"

			' Set the font size of the text to 80
			txtRang.CharacterFormat.FontSize = 80

			' Set the text color of the text to SeaGreen
			txtRang.CharacterFormat.TextColor = Color.SeaGreen

			' Specify the output file name as "CreateBarcode.docx"
			Dim output As String = "CreateBarcode.docx"

			' Save the document to the specified file in Docx format
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose the Document object to release resources
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
