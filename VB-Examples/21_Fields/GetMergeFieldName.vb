Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text

Namespace GetMergeFieldName
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a StringBuilder to hold the field information
			Dim sb As New StringBuilder()

			' Load the document from a file
			Dim document As New Document("..\..\..\..\..\..\Data\MailMerge.doc")

			' Get the array of merge field names in the document
			Dim fieldNames() As String = document.MailMerge.GetMergeFieldNames()

			' Append the count of merge fields in the document to the StringBuilder
			sb.Append("The document has " & fieldNames.Length.ToString() & " merge fields.")

			' Append a header for the merge field names
			sb.Append(" The below is the name of the merge field:" & vbCrLf)

			' Iterate through each merge field name and append it to the StringBuilder
			For Each name As String In fieldNames
				sb.AppendLine(name)
			Next name

			' Write the result to a text file
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
