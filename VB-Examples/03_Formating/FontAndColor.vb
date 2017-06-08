Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace FontAndColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            'Open a blank word document as template
            Dim document_Renamed As New Document("..\..\..\..\..\..\Data\Blank.doc")

            'Get the first secition
            Dim section As Section = document_Renamed.Sections(0)

            'Create a new paragraph or get the first paragraph
            Dim paragraph_Renamed As Paragraph = Nothing
            If section.Paragraphs.Count > 0 Then
                paragraph_Renamed = section.Paragraphs(0)
            Else
                paragraph_Renamed = section.AddParagraph()
            End If

            'Append Text
            Dim text As String _
                = "This paragraph is demo of text font and color. " _
                & "The font name of this paragraph is Tahoma. " _
                & "The font size of this paragraph is 20. " _
                & "The under line style of this paragraph is DotDot. " _
                & "The color of this paragraph is Blue. "
            Dim txtRange As TextRange = paragraph_Renamed.AppendText(text)

            'Font name
            txtRange.CharacterFormat.FontName = "Tahoma"

            'Font size
            txtRange.CharacterFormat.FontSize = 20

            'Underline
            txtRange.CharacterFormat.UnderlineStyle = UnderlineStyle.DotDot

            'Change text color
            txtRange.CharacterFormat.TextColor = Color.Blue

			'Save doc file.
			document_Renamed.SaveToFile("Sample.doc",FileFormat.Doc)

			'Launching the MS Word file.
			WordDocViewer("Sample.doc")


		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
