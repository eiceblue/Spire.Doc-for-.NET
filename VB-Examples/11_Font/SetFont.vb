Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports Spire.Doc.Fields

Namespace SetFont
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\Sample.docx"

			'Create a Word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first section 
			Dim s As Section = doc.Sections(0)

			'Get the second paragraph
			Dim p As Paragraph = s.Paragraphs(1)

			'Create a characterFormat object
			Dim format As New CharacterFormat(doc)
			'Set font
			format.Font = New Font("Arial", 16)
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'format.FontName = "Arial";
			'format.FontSize = 16;
			' =============================================================================

			'Loop through the childObjects of paragraph 
			For Each childObj As DocumentObject In p.ChildObjects
				If TypeOf childObj Is TextRange Then
					'Apply character format
					Dim tr As TextRange = TryCast(childObj, TextRange)
					tr.ApplyCharacterFormat(format)
				End If
			Next childObj

			'Save the document
			Dim output As String = "SetFont.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			'Dispose the document.
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
