Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace GetParagraphRevisionsDetails
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\Data\Revisions.docx")

			' Create a StringBuilder object to store the output
			Dim builder As New StringBuilder()

			' Iterate through each Section in the document
			For Each section As Section In document.Sections

				' Iterate through each Paragraph in the current Section
				For Each paragraph As Paragraph In section.Paragraphs
				
					' Check if the current Paragraph is a deleted revision
					If paragraph.IsDeleteRevision Then
					
						' Append information about the deleted revision to the output
						builder.AppendLine(String.Format("The section {0} paragraph {1} has been changed (deleted).", document.GetIndex(section), section.GetIndex(paragraph)))
						builder.AppendLine("Author: " & paragraph.DeleteRevision.Author)
						builder.AppendLine("DateTime: " & paragraph.DeleteRevision.DateTime)
						builder.AppendLine("Type: " & paragraph.DeleteRevision.Type)
						builder.AppendLine("")
						
					' Check if the current Paragraph is an inserted revision
					ElseIf paragraph.IsInsertRevision Then
					
						' Append information about the inserted revision to the output
						builder.AppendLine(String.Format("The section {0} paragraph {1} has been changed (inserted).", document.GetIndex(section), section.GetIndex(paragraph)))
						builder.AppendLine("Author: " & paragraph.InsertRevision.Author)
						builder.AppendLine("DateTime: " & paragraph.InsertRevision.DateTime)
						builder.AppendLine("Type: " & paragraph.InsertRevision.Type)
						builder.AppendLine("")
						
					Else
					
						' Iterate through each ChildObject in the current Paragraph
						For Each obj As DocumentObject In paragraph.ChildObjects
						
							' Check if the current ChildObject is a TextRange
							If obj.DocumentObjectType.Equals(DocumentObjectType.TextRange) Then
							
								' Cast the ChildObject to a TextRange
								Dim textRange As TextRange = TryCast(obj, TextRange)
								
								' Check if the current TextRange is a deleted revision
								If textRange.IsDeleteRevision Then
								
									' Append information about the deleted revision to the output
									builder.AppendLine(String.Format("The section {0} paragraph {1} textrange {2} has been changed (deleted).", document.GetIndex(section), section.GetIndex(paragraph), paragraph.GetIndex(textRange)))
									builder.AppendLine("Author: " & textRange.DeleteRevision.Author)
									builder.AppendLine("DateTime: " & textRange.DeleteRevision.DateTime)
									builder.AppendLine("Type: " & textRange.DeleteRevision.Type)
									builder.AppendLine("Change Text: " & textRange.Text)
									builder.AppendLine("")
									
								' Check if the current TextRange is an inserted revision
								ElseIf textRange.IsInsertRevision Then
								
									' Append information about the inserted revision to the output
									builder.AppendLine(String.Format("The section {0} paragraph {1} textrange {2} has been changed (inserted).", document.GetIndex(section), section.GetIndex(paragraph), paragraph.GetIndex(textRange)))
									builder.AppendLine("Author: " & textRange.InsertRevision.Author)
									builder.AppendLine("DateTime: " & textRange.InsertRevision.DateTime)
									builder.AppendLine("Type: " & textRange.InsertRevision.Type)
									builder.AppendLine("Change Text: " & textRange.Text)
									builder.AppendLine("")
									
								End If
							End If
						Next obj
					End If
				Next paragraph

			' Specify the filename for the output result
			Dim output As String = "GetParagraphRevisionsDetails.txt"

			' Write the content to the specified file
			File.WriteAllText(output, builder.ToString())

			'Launch the file
			TxtViewer(output)
			Next section
			' Dispose the Document object to release resources
			document.Dispose()
		End Sub

		Private Sub TxtViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub


	End Class
End Namespace
