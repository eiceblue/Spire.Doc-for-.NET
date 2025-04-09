Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace GetListOfUsingFonts
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\UsingFonts.docx"
			Dim output As String = "GetListOfUsingFonts.txt"
			Dim stringBuilder As New StringBuilder()
			Dim font_obj As New Dictionary(Of Font, TextRange)() From {}

			'Create a Word document
			Dim document As New Document()

			'Load the file from disk
			document.LoadFromFile(input)

			'Loop through the sections,paragraphs and child objects of paragraph
			For Each section As Section In document.Sections
				For Each paragraph As Paragraph In section.Body.Paragraphs
					For Each obj As DocumentObject In paragraph.ChildObjects
						If obj.DocumentObjectType.Equals(DocumentObjectType.TextRange) Then
							'Get the text range
							Dim range As TextRange = TryCast(obj, TextRange)

							'Get the font info
							Dim font As Font = range.CharacterFormat.Font
							If Not font_obj.ContainsKey(font) Then

								'Add the font to the dictionary
								font_obj.Add(font, range)
							End If
						End If
					Next obj
				Next paragraph
			Next section

			'Loop through the dictionary 
			For Each item In font_obj

				'Get the font properties
				Dim font As Font = item.Key
				Dim range As TextRange = item.Value
				Dim s As String = String.Format("Font Name: {0}, Size:{1}, Style:{2}, Color:{3}", font.Name, font.Size, font.Style, range.CharacterFormat.TextColor.Name)
				stringBuilder.AppendLine(s)
			Next item

			File.WriteAllText(output, stringBuilder.ToString())

			'Dispose the document.
			document.Dispose()

			'Launching the Text file.
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
