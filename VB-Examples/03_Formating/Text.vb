Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace Text
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
            paragraph_Renamed.AppendText("The various ways to format paragraph text in Microsoft Word:")

            paragraph_Renamed.ApplyStyle(BuiltinStyle.Heading1)

            'Append alignment text
            AppendAligmentText(section)

            'Append indentation text
            AppendIndentationText(section)

            AppendBulletedList(section)

            'Save doc file.
            document_Renamed.SaveToFile("Sample.doc", FileFormat.Doc)

            'Launching the MS Word file.
            WordDocViewer("Sample.doc")


        End Sub

		Private Sub AppendAligmentText(ByVal section_Renamed As Section)
			Dim paragraph_Renamed As Paragraph = Nothing

			paragraph_Renamed = section_Renamed.AddParagraph()

			'Append Text
			paragraph_Renamed.AppendText("Horizontal Aligenment")

			paragraph_Renamed.ApplyStyle(BuiltinStyle.Heading3)

			For Each align As Spire.Doc.Documents.HorizontalAlignment In System.Enum.GetValues(GetType(Spire.Doc.Documents.HorizontalAlignment))
				Dim paramgraph As Paragraph = section_Renamed.AddParagraph()
				paramgraph.AppendText("This text is " & align.ToString())
				paramgraph.Format.HorizontalAlignment = align
			Next align
		End Sub

		Private Sub AppendIndentationText(ByVal section_Renamed As Section)
			Dim paragraph_Renamed As Paragraph = Nothing

			paragraph_Renamed = section_Renamed.AddParagraph()

			'Append Text
			paragraph_Renamed.AppendText("Indentation")

			paragraph_Renamed.ApplyStyle(BuiltinStyle.Heading3)

			paragraph_Renamed = section_Renamed.AddParagraph()
			paragraph_Renamed.AppendText("Indentation is the spacing between text and margins. Word allows you to set left and right margins, as well as indentations for the first line of a paragraph and hanging indents")
			paragraph_Renamed.Format.FirstLineIndent = 15
		End Sub

		Private Sub AppendBulletedList(ByVal section_Renamed As Section)
			Dim paragraph_Renamed As Paragraph = Nothing

			paragraph_Renamed = section_Renamed.AddParagraph()


			'Append Text
			paragraph_Renamed.AppendText("Bulleted List")

			paragraph_Renamed.ApplyStyle(BuiltinStyle.Heading3)

			paragraph_Renamed = section_Renamed.AddParagraph()
			For i As Integer = 0 To 4
				paragraph_Renamed = section_Renamed.AddParagraph()
				paragraph_Renamed.AppendText("Item" & i.ToString())

				If i = 0 Then
					paragraph_Renamed.ListFormat.ApplyBulletStyle()
				Else
					paragraph_Renamed.ListFormat.ContinueListNumbering()
				End If

				paragraph_Renamed.ListFormat.ListLevelNumber = 1
			Next i
		End Sub





		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
