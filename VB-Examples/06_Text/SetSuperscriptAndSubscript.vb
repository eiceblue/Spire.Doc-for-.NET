Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetSuperscriptAndSubscript
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object.
			Dim document As New Document()

			' Add a section to the document.
			Dim section As Section = document.AddSection()

			' Add a paragraph to the section.
			Dim paragraph As Paragraph = section.AddParagraph()
			paragraph.AppendText("E = mc")

			' Append the superscript "2" to the paragraph and store it in range1.
			Dim range1 As TextRange = paragraph.AppendText("2")
			range1.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript

			' Insert a line break in the paragraph.
			paragraph.AppendBreak(BreakType.LineBreak)

			' Append the text "F" to the paragraph and store it in range2.
			paragraph.AppendText("F")
			Dim range2 As TextRange = paragraph.AppendText("n")
			range2.CharacterFormat.SubSuperScript = SubSuperScript.SubScript

			' Append more text to the paragraph with subscripts.
			paragraph.AppendText(" = F")
			paragraph.AppendText("n-1").CharacterFormat.SubSuperScript = SubSuperScript.SubScript
			paragraph.AppendText(" + F")
			paragraph.AppendText("n-2").CharacterFormat.SubSuperScript = SubSuperScript.SubScript

			' Iterate through each item in the paragraph and set the font size to 36 for TextRange objects.
			For Each i In paragraph.Items
				If TypeOf i Is TextRange Then
					TryCast(i, TextRange).CharacterFormat.FontSize = 36
				End If
			Next i

			' Save the document to a file named "SetSuperscriptAndSubscript.docx".
			Dim output As String = "SetSuperscriptAndSubscript.docx"
			document.SaveToFile(output, FileFormat.Docx)

			' Dispose of the document object to free up resources.
			document.Dispose()

			'Launching the file
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
