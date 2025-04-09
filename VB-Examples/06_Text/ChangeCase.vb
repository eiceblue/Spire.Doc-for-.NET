Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ChangeCase
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Specify the input file location for the Word document.
			Dim input As String = "..\..\..\..\..\..\Data\Text1.docx"

			'Create a new Document object.
			Dim doc As New Document()

			'Load the Word document from the specified input file location.
			doc.LoadFromFile(input)

			'Declare a TextRange variable.
			Dim textRange As TextRange

			'Access the first paragraph in the first section of the document.
			Dim para1 As Paragraph = doc.Sections(0).Paragraphs(1)

			'Iterate through each child object in the first paragraph.
			For Each obj As DocumentObject In para1.ChildObjects
				'Check if the child object is of type TextRange.
				If TypeOf obj Is TextRange Then
					'Assign the child object to the TextRange variable.
					textRange = TryCast(obj, TextRange)
					'Set the AllCaps property of the CharacterFormat object associated with the text range to True.
					textRange.CharacterFormat.AllCaps = True
				End If
			Next obj

			'Access the third paragraph in the first section of the document.
			Dim para2 As Paragraph = doc.Sections(0).Paragraphs(3)

			'Iterate through each child object in the second paragraph.
			For Each obj As DocumentObject In para2.ChildObjects
				'Check if the child object is of type TextRange.
				If TypeOf obj Is TextRange Then
					'Assign the child object to the TextRange variable.
					textRange = TryCast(obj, TextRange)
					'Set the IsSmallCaps property of the CharacterFormat object associated with the text range to True.
					textRange.CharacterFormat.IsSmallCaps = True
				End If
			Next obj

			'Specify the output file name for the modified document.
			Dim output As String = "ChangeCase.docx"

			'Save the modified document to the specified output file location using the Docx2013 file format.
			doc.SaveToFile(output, FileFormat.Docx2013)

			'Dispose of the Document object to release system resources.
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
