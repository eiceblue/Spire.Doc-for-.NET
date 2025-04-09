Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text

Namespace GetFieldText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a StringBuilder to hold the field information
			Dim sb As New StringBuilder()

			' Load the document from a file
			Dim document As New Document("..\..\..\..\..\..\Data\SampleB_1.docx")

			' Get the collection of fields in the document
			Dim fields As FieldCollection = document.Fields

			' Iterate through each field in the collection
			For Each field As Field In fields
				' Get the text of the field
				Dim fieldText As String = field.FieldText

				' Append the field text to the StringBuilder
				sb.Append("The field text is """ & fieldText & """." & vbCrLf)
			Next field

			' Write the field information to a text file
			File.WriteAllText("result.txt", sb.ToString())

			' Dispose the document object
			document.Dispose()

			'Launch result file
			WordDocViewer("result.txt")

		End Sub


		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
