Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting.Revisions
Imports Spire.Doc.Fields
Imports System.IO

Namespace GetMoveRevisions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            ' Load the existing Word document that contains tracked moves
            Dim document As Document = New Document("..\..\..\..\..\..\..\Data\MoveRevision.docx")

            ' Create a DifferRevisions object to access move revisions within the document
            Dim differRevisions As DifferRevisions = New DifferRevisions(document)

            ' Get the list of 'Move From' revisions (content that was moved from one location)
            Dim moveFromRevisions As List(Of DocumentObject) = differRevisions.MoveFromRevisions

            ' Get the list of 'Move To' revisions (content that was moved to a new location)
            Dim moveToRevisions As List(Of DocumentObject) = differRevisions.MoveToRevisions

            ' Create a StringBuilder to accumulate information about 'Move From' revisions
            Dim moveFromRevisions_content As StringBuilder = New StringBuilder()

            ' Append a header line indicating the count of 'Move From' revisions
            moveFromRevisions_content.AppendLine("MoveFromRevisions: " & moveFromRevisions.Count)

            ' Loop through each 'Move From' revision object
            For i As Integer = 0 To moveFromRevisions.Count - 1
                ' Append the string representation of the revision object
                moveFromRevisions_content.AppendLine(moveFromRevisions(i).ToString())

                ' Check if the revision object is a Paragraph
                If moveFromRevisions(i).DocumentObjectType = DocumentObjectType.Paragraph Then
                    ' If it's a paragraph, append its text content
                    moveFromRevisions_content.AppendLine(DirectCast(moveFromRevisions(i), Paragraph).Text)
                End If

                ' Check if the revision object is a TextRange (a piece of text)
                If moveFromRevisions(i).DocumentObjectType = DocumentObjectType.TextRange Then
                    ' If it's a text range, append its text content
                    moveFromRevisions_content.AppendLine(DirectCast(moveFromRevisions(i), TextRange).Text)
                End If
            Next

            ' Create a StringBuilder to accumulate information about 'Move To' revisions
            Dim moveToRevisions_content As StringBuilder = New StringBuilder()

            ' Append a header line indicating the count of 'Move To' revisions
            moveToRevisions_content.AppendLine("MoveToRevisions: " & moveToRevisions.Count)

            ' Loop through each 'Move To' revision object
            For i As Integer = 0 To moveToRevisions.Count - 1
                ' Append the string representation of the revision object
                moveToRevisions_content.AppendLine(moveToRevisions(i).ToString())

                ' Check if the revision object is a Paragraph
                If moveToRevisions(i).DocumentObjectType = DocumentObjectType.Paragraph Then
                    ' If it's a paragraph, append its text content
                    moveToRevisions_content.AppendLine(DirectCast(moveToRevisions(i), Paragraph).Text)
                End If

                ' Check if the revision object is a TextRange (a piece of text)
                If moveToRevisions(i).DocumentObjectType = DocumentObjectType.TextRange Then
                    ' If it's a text range, append its text content
                    moveToRevisions_content.AppendLine(DirectCast(moveToRevisions(i), TextRange).Text)
                End If
            Next

            ' Write the accumulated 'Move From' revision information to a text file
            File.WriteAllText("MoveFromRevisions.txt", moveFromRevisions_content.ToString())

            ' Write the accumulated 'Move To' revision information to a text file
            File.WriteAllText("MoveToRevisions.txt", moveToRevisions_content.ToString())

            ' Dispose of the document object to free resources
            document.Dispose()

        End Sub
	End Class
End Namespace
