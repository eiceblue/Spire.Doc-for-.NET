Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ModifyHyperlinkText
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

		   ' Create a list to store the hyperlinks
		   Dim hyperlinks As New List(Of Field)()

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
							   End If
						   End If
					   Next para
				   End If
			   Next sec
		   Next section

		   ' Modify the text of the first hyperlink field
		   hyperlinks(0).FieldText = "Spire.Doc component"

		   ' Specify the output file path for the modified document
		   Dim output As String = "ModifyText.docx"

		   ' Save the modified document to the output file path in DOCX format
		   doc.SaveToFile(output, FileFormat.Docx)

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
