Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertNewText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim input As String = "..\..\..\..\..\..\Data\Sample.docx"

			' Create a new Document object
			Dim doc As New Document()

			' Load the document from the specified input file
			doc.LoadFromFile(input)

			' Find all occurrences of the word "Word" in the document and store them in an array of TextSelection objects
			Dim selections() As TextSelection = doc.FindAllString("Word", True, True)

			' Initialize variables
			Dim index As Integer = 0
			Dim range As New TextRange(doc)

			' Iterate through each TextSelection in the selections array
			For Each selection As TextSelection In selections
				' Get the selected text range as one complete range
				range = selection.GetAsOneRange()

				' Create a new TextRange object
				Dim newrange As New TextRange(doc)

				' Set the text of the new range to "(New text)"
				newrange.Text = "(New text)"

				' Find the index of the range within its owner paragraph's ChildObjects collection
				index = range.OwnerParagraph.ChildObjects.IndexOf(range)

				' Insert the new range after the current range in the owner paragraph's ChildObjects collection
				range.OwnerParagraph.ChildObjects.Insert(index + 1, newrange)
			Next selection

			' Find all occurrences of the text "New text" in the document and store them in an array of TextSelection objects
			Dim text() As TextSelection = doc.FindAllString("New text", True, True)

			' Iterate through each TextSelection in the text array
			For Each seletion As TextSelection In text
				' Get the selected text range as one complete range and set its CharacterFormat's HighlightColor to Yellow
				seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow
			Next seletion

			' Specify the output file path
			Dim output As String = "InsertNewText.docx"

			' Save the modified document to the specified output file in DOCX format
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose of the Document object to release resources
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
