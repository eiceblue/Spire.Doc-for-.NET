Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace GetCharacterSpacing
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a document
			Dim document As New Document()

			'Load the document from disk.
			document.LoadFromFile("..\..\..\..\..\..\Data\Insert.docx")

			'Get the first section of document
			Dim section As Section = document.Sections(0)

			'Get the first paragraph 
			Dim paragraph As Paragraph = section.Paragraphs(0)

			'Define two variables
			Dim fontName As String = ""
			Dim fontSpacing As Single = 0

			'Traverse the ChildObjects 
			For Each docObj As DocumentObject In paragraph.ChildObjects
				'If it is TextRange
				If TypeOf docObj Is TextRange Then
					Dim textRange As TextRange = TryCast(docObj, TextRange)

					'Get the font name
					fontName = textRange.CharacterFormat.Font.Name
					' =============================================================================
					' Use the following code for netstandard dlls
					' =============================================================================
					' fontName = textRange.CharacterFormat.FontName
					' =============================================================================

					'Get the character spacing
					fontSpacing = textRange.CharacterFormat.CharacterSpacing
				End If
			Next docObj

			'Show the result in message box
			MessageBox.Show("The font of first paragraph is " & fontName & ", the character spacing is " & fontSpacing & "pt.")
			' Dispose of the document object
			document.Dispose()
			
		End Sub
	End Class
End Namespace
