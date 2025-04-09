Imports Spire.Doc

Namespace SetFontFallbackRule
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
'            Instructions:
'             Support for switching fonts that do not support drawing characters through the FontFallbackRule method in XML when converting to a non-flow layout document.
'
'             If there is no XML available, first save an XML using saveFontFallbackRuleSettings and then manually edit the font replacement rules in the XML.
'             The rules consist of three attributes: Ranges correspond to Unicode ranges for each character; FallbackFonts correspond to the font names for substitution; BaseFonts correspond to the font names for characters in the document.
'             When editing the XML, it is important to note that the rules are searched from top to bottom for character matching.
'             After editing the XML, load the rules using the loadFontFallbackRuleSettings method.
'             

			' Create a new Document object
			Dim doc As New Document()

			' Load the document from the specified file
			doc.LoadFromFile("..\..\..\..\..\..\..\Data\SetFontFallbackRule.docx")

			' Save the font fallback rule settings to an XML file
			'doc.SaveFontFallbackRuleSettings("fontSettings.xml");

			' Load the font fallback rule settings from the XML file
			doc.LoadFontFallbackRuleSettings("..\..\..\..\..\..\..\Data\FontFallbackRule.xml")

			' Save the document to a PDF file with the specified output file name
			doc.SaveToFile("SetFontFallbackRule_output.pdf", FileFormat.PDF)

			' Dispose the document object
			doc.Dispose()

			'Launch result file
			WordDocViewer("SetFontFallbackRule_output.pdf")

		End Sub


		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
