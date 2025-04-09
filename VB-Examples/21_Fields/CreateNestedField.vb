Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Interface

Namespace CreateNestedField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a file
			Dim document As New Document("..\..\..\..\..\..\Data\SampleB_2.docx")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Add a paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Create the outer IF field and add it to the paragraph
			Dim ifField As New IfField(document)
			ifField.Type = FieldType.FieldIf
			ifField.Code = "IF "
			paragraph.Items.Add(ifField)

			' Create the inner IF field and add it to the paragraph
			Dim ifField2 As New IfField(document)
			ifField2.Type = FieldType.FieldIf
			ifField2.Code = "IF "
			paragraph.ChildObjects.Add(ifField2)
			paragraph.Items.Add(ifField2)
			paragraph.AppendText("""200"" < ""50""   ""200"" ""50"" ")

			' Create the end mark for the inner IF field and add it to the paragraph
			Dim embeddedEnd As IParagraphBase = document.CreateParagraphItem(ParagraphItemType.FieldMark)
			TryCast(embeddedEnd, FieldMark).Type = FieldMarkType.FieldEnd
			paragraph.Items.Add(embeddedEnd)
			ifField2.End = TryCast(embeddedEnd, FieldMark)

			' Append additional text and create the end mark for the outer IF field
			paragraph.AppendText(" > ")
			paragraph.AppendText("""100"" ")
			paragraph.AppendText("""Thanks"" ")
			paragraph.AppendText("""The minimum order is 100 units""")
			Dim [end] As IParagraphBase = document.CreateParagraphItem(ParagraphItemType.FieldMark)
			TryCast([end], FieldMark).Type = FieldMarkType.FieldEnd
			paragraph.Items.Add([end])
			ifField.End = TryCast([end], FieldMark)

			' Enable field update
			document.IsUpdateFields = True

			' Specify the file name for saving the document
			Dim result As String = "CreateNestedField_output.docx"

			' Save the document to a file
			document.SaveToFile(result, FileFormat.Docx)

			' Dispose the document object
			document.Dispose()
			
			'Launch result file
			WordDocViewer(result)

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
