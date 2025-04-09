Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace RemoveEditableRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
		   ' Create a new Document object
		   Dim document As New Document()

		   ' Load the document from the specified file path
		   document.LoadFromFile("..\..\..\..\..\..\Data\RemoveEditableRange.docx")

		   ' Iterate through each section in the document
		   For Each section As Section In document.Sections
			   ' Iterate through each paragraph in the section's body
			   For Each paragraph As Paragraph In section.Body.Paragraphs
				   ' Loop through the child objects of the paragraph
				   Dim i As Integer = 0
				   Do While i < paragraph.ChildObjects.Count
					   Dim obj As DocumentObject = paragraph.ChildObjects(i)

					   ' Check if the child object is a PermissionStart or PermissionEnd element
					   If TypeOf obj Is PermissionStart OrElse TypeOf obj Is PermissionEnd Then
						   ' Remove the PermissionStart or PermissionEnd element from the paragraph
						   paragraph.ChildObjects.Remove(obj)
					   Else
						   ' Move to the next child object
						   i += 1
					   End If
				   Loop
			   Next paragraph
		   Next section

		   ' Specify the output file path for the modified document
		   Dim output As String = "RemoveEditableRange_output.docx"

		   ' Save the modified document to the output file path in DOCX format
		   document.SaveToFile(output, FileFormat.Docx)

		   ' Dispose the document object to free up resources
		   document.Dispose()
			
			WordDocViewer(output)
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
