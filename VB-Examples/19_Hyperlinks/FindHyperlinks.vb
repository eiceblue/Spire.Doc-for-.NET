Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace FindHyperlinks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
		
		   ' Specify the input file path for the document containing hyperlinks
		   Dim input As String = "..\..\..\..\..\..\Data\Hyperlinks.docx"

		   ' Create a new Document object
		   Dim doc As New Document()

		   ' Load the document from the specified file path
		   doc.LoadFromFile(input)

		   ' Create a list to store the hyperlinks and a variable to hold the text of the hyperlinks
		   Dim hyperlinks As New List(Of Field)()
		   Dim hyperlinksText As String = Nothing

		   ' Iterate through the sections in the document
		   For Each section As Section In doc.Sections
			   ' Iterate through the child objects in the body of the section
			   For Each sec As DocumentObject In section.Body.ChildObjects
				   ' Check if the child object is a paragraph
				   If sec.DocumentObjectType = DocumentObjectType.Paragraph Then
					   ' Iterate through the child objects in the paragraph
					   For Each para As DocumentObject In (TryCast(sec, Paragraph)).ChildObjects
						   ' Check if the child object is a field
						   If para.DocumentObjectType = DocumentObjectType.Field Then
							   ' Cast the child object to a Field
							   Dim field As Field = TryCast(para, Field)

							   ' Check if the field is a hyperlink
							   If field.Type = FieldType.FieldHyperlink Then
								   ' Add the field to the list of hyperlinks
								   hyperlinks.Add(field)

								   ' Append the field's text to the hyperlinksText variable
								   hyperlinksText &= field.FieldText & vbCrLf
							   End If
						   End If
					   Next para
				   End If
			   Next sec
		   Next section

		   ' Specify the output file path for the generated text file
		   Dim output As String = "HyperlinksText.txt"

		   ' Write the hyperlinks text to the output file
		   File.WriteAllText(output, hyperlinksText)

		   ' Dispose the document object to free up resources
		   doc.Dispose()
			
			Viewer(output)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
