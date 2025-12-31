Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ReplaceWithImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path for the Word document.
			Dim input As String = "..\..\..\..\..\..\Data\Template.docx"

			' Create a new Document object.
			Dim doc As New Document()

			' Load the Word document from the specified input file.
			doc.LoadFromFile(input)
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim selections() As TextSelection = doc.FindAllString("E-iceblue", True, True)
			'Dim index As Integer = 0
			'Dim range As TextRange = Nothing

			'For Each selection As TextSelection In selections
			'	Dim pic As New DocPicture(doc)
			'	pic.LoadImage(inputFile_2)

			'	range = selection.GetAsOneRange()
			'	index = range.OwnerParagraph.ChildObjects.IndexOf(range)
			'	range.OwnerParagraph.ChildObjects.Insert(index, pic)
			'	range.OwnerParagraph.ChildObjects.Remove(range)
			'Next
			' =============================================================================


			' Create an Image object from the image file.
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\E-iceblue.png")

			' Find all occurrences of the text "E-iceblue" in the document.
			Dim selections() As TextSelection = doc.FindAllString("E-iceblue", True, True)

			' Variables for storing the index and range of the text to be replaced.
			Dim index As Integer = 0
			Dim range As TextRange = Nothing

			' Iterate through each text selection.
			For Each selection As TextSelection In selections
				' Create a new DocPicture object and load the image into it.
				Dim pic As New DocPicture(doc)
				pic.LoadImage(image)

				' Get the text range as a single range and get its index within the owner paragraph.
				range = selection.GetAsOneRange()
				index = range.OwnerParagraph.ChildObjects.IndexOf(range)

				' Insert the picture at the same index within the owner paragraph,
				' and remove the original text range.
				range.OwnerParagraph.ChildObjects.Insert(index, pic)
				range.OwnerParagraph.ChildObjects.Remove(range)
			Next selection

			' Specify the output file name for saving the modified document.
			Dim output As String = "ReplaceWithImage.docx"

			' Save the modified document to the specified output file in Docx format.
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose of the Document object to release resources.
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
