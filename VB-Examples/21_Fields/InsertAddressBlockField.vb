Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text

Namespace InsertAddressBlockField
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

			' Add a new paragraph to the section
			Dim par As Paragraph = section.AddParagraph()

			' Append a field with type "AddressBlock" to the paragraph
			Dim field As Field = par.AppendField("ADDRESSBLOCK", FieldType.FieldAddressBlock)

			' Set the code for the field, including additional options and formatting
			field.Code = "ADDRESSBLOCK \c 1 \d \e Test2 \f Test3 \l ""Test 4"""

			' Save the modified document to a file
			document.SaveToFile("result.docx", FileFormat.Docx)

			' Dispose the document object
			document.Dispose() 
			
			'Launch result file
			WordDocViewer("result.docx")

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
