Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddCheckBoxContentControl
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
			Dim txtRange As TextRange = paragraph.AppendText("The following example shows how to add CheckBox content control in a Word document. " & vbLf)

			' Append text indicating adding the CheckBox content control
			txtRange = paragraph.AppendText("Add CheckBox Content Control:  ")

			' Set the text range formatting to italic
			txtRange.CharacterFormat.Italic = True

			' Create an inline structure document tag (SDT) and add it to the paragraph's child objects
			Dim sdt As New StructureDocumentTagInline(document)
			paragraph.ChildObjects.Add(sdt)

			' Set the SDT type to CheckBox
			sdt.SDTProperties.SDTType = SdtType.CheckBox

			' Create an instance of SdtCheckBox and set it as the control properties for the SDT
			Dim scb As New SdtCheckBox()
			sdt.SDTProperties.ControlProperties = scb

			' Create a TextRange object, set its font name and size
			Dim tr As New TextRange(document)
			tr.CharacterFormat.FontName = "MS Gothic"
			tr.CharacterFormat.FontSize = 12

			' Add the TextRange object to the SDT's child objects
			sdt.ChildObjects.Add(tr)

			' Set the CheckBox as checked
			scb.Checked = True

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
