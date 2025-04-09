Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.IO

Namespace IdentifyMergeFieldName
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create Word document.
			Dim document As New Document()

			'Load the file from disk.
			document.LoadFromFile("..\..\..\..\..\..\Data\IdentifyMergeFieldNames.docx")

			'Get the collection of group names.
			Dim GroupNames() As String = document.MailMerge.GetMergeGroupNames()

			'Get the collection of merge field names in a specific group.
			Dim MergeFieldNamesWithinRegion() As String = document.MailMerge.GetMergeFieldNames("Products")

			'Get the collection of all the merge field names.
			Dim MergeFieldNames() As String = document.MailMerge.GetMergeFieldNames()

			Dim content As New StringBuilder()
			content.AppendLine("----------------Group Names-----------------------------------------")
			For i As Integer = 0 To GroupNames.Length - 1
				content.AppendLine(GroupNames(i))
			Next i

			content.AppendLine("----------------Merge field names within a specific group-----------")
			For j As Integer = 0 To MergeFieldNamesWithinRegion.Length - 1
				content.AppendLine(MergeFieldNamesWithinRegion(j))
			Next j

			content.AppendLine("----------------All of the merge field names------------------------")
			For k As Integer = 0 To MergeFieldNames.Length - 1
				content.AppendLine(MergeFieldNames(k))
			Next k

			Dim result As String = "Result-IdentifyMergeFieldNames.txt"

			'Save to file.
			File.WriteAllText(result,content.ToString())
			
			'Dispose the document
			document.Dispose()
			
			'Launch the file.
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
