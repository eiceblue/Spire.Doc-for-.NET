Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace Styles
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
			paragraph_Renamed.AppendText("Builtin Style:")

			For Each builtinStyle_Renamed As BuiltinStyle In System.Enum.GetValues(GetType(BuiltinStyle))
                paragraph_Renamed = section.AddParagraph()
				'Append Text
				paragraph_Renamed.AppendText(builtinStyle_Renamed.ToString())
				'Apply Style
				paragraph_Renamed.ApplyStyle(builtinStyle_Renamed)
			Next builtinStyle_Renamed

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
