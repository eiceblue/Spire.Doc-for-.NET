Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace InsertField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new DocumentNavigator instance, which automatically creates an underlying empty document.
			Dim navigator As New DocumentNavigator()

			' Get the Document object associated with the navigator.
			Dim doc As Document = navigator.Document

			' Retrieve the current (and initially only) section of the document.
			Dim section As Section = navigator.CurrentSection

			' Insert a new paragraph at the current position and assign it to 'par'.
			Dim par As Paragraph = navigator.InsertParagraph()

			' Write the label text "Add page fields:" followed by a paragraph break.
			navigator.Writeln("Add page fields:")

			' Insert a PAGE field with numeric formatting.
			navigator.InsertField("PAGE \# ""Page 0""")

			' Insert a paragraph break (empty line).
			navigator.Writeln()

			' Insert the same PAGE field but with a custom result text to simulate a placeholder or initial value.
			navigator.InsertField("PAGE \# ""Page 0""", "3")

			' Insert another paragraph break.
			navigator.Writeln()

			' Insert a built-in PAGE field using FieldType enumeration, with the field result displayed (true = show result).
			navigator.InsertField(FieldType.FieldPage, True)

			' Insert another paragraph break.
			navigator.Writeln()

			' Insert another PAGE field, but this time hide the field result (false = show field code instead of result).
			navigator.InsertField(FieldType.FieldPage, False)

			' Save the document to a file in DOCX format.
			doc.SaveToFile("InsertField.docx", FileFormat.Docx)


			' Close the document to release internal resources.
			doc.Close()

			' Explicitly dispose of the document object to free memory immediately.
			doc.Dispose()

			'Launch the MS Word file.
			WordDocViewer("InsertField.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
