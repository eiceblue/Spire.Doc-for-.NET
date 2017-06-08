Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace Indent
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
            paragraph_Renamed.AppendText("Using items list to show Indent demo.")

            paragraph_Renamed.ApplyStyle(BuiltinStyle.Heading3)

            paragraph_Renamed = section.AddParagraph()
            For i As Integer = 0 To 9
                paragraph_Renamed = section.AddParagraph()
                Dim text As String _
                    = "Indent Demo Node" & i.ToString()
                Dim txtRange As TextRange = paragraph_Renamed.AppendText(text)


                If i = 0 Then
                    paragraph_Renamed.ListFormat.ApplyBulletStyle()
                Else
                    paragraph_Renamed.ListFormat.ContinueListNumbering()

                End If

                paragraph_Renamed.ListFormat.CurrentListLevel.NumberPosition = -10

            Next i

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
