Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddRichTextContentControl
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim document As New Document()

			' Add a section to the document
			Dim section As Section = document.AddSection()

			' Add a paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append text with an explanation to the paragraph
			Dim txtRange As TextRange = paragraph.AppendText("The following example shows how to add RichText content control in a Word document. " & vbLf)

			' Append text indicating adding the RichText content control
			txtRange = paragraph.AppendText("Add RichText Content Control:  ")

			' Set the text range formatting to italic
			txtRange.CharacterFormat.Italic = True

			' Create an inline structure document tag (SDT) and add it to the paragraph's child objects
			Dim sdt As New StructureDocumentTagInline(document)
			paragraph.ChildObjects.Add(sdt)

			' Set the SDT type to RichText
			sdt.SDTProperties.SDTType = SdtType.RichText

			' Create an instance of SdtText, set its multiline property, and assign it as the control properties for the SDT
			Dim text As New SdtText(True)
			text.IsMultiline = True
			sdt.SDTProperties.ControlProperties = text

			' Create a TextRange object and set its text and text color, then add it to the SDT's content
			Dim rt As New TextRange(document)
			rt.Text = "Welcome to use "
			rt.CharacterFormat.TextColor = Color.Green
			sdt.SDTContent.ChildObjects.Add(rt)

			' Create another TextRange object and set its text and text color, then add it to the SDT's content
			rt = New TextRange(document)
			rt.Text = "Spire.Doc"
			rt.CharacterFormat.TextColor = Color.OrangeRed
			sdt.SDTContent.ChildObjects.Add(rt)

			' Save the document to a file in Docx format
			document.SaveToFile("Output.docx", FileFormat.Docx)

			' Dispose the document object
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
