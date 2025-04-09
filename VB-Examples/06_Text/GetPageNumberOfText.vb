Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports System.IO
Imports Spire.Doc.Documents
Imports Spire.Doc.Pages

Namespace GetPageNumberOfText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object.
			Dim document As New Document()

			'Load a Word document from a file.
			document.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			'Find all occurrences of the word "Spire" in the document and store them in an array of TextSelection objects.
			Dim textSelections() As TextSelection = document.FindAllString("Spire", False, False)

			'Create a new FixedLayoutDocument object using the loaded document.
			Dim layoutDoc As New FixedLayoutDocument(document)

			'Initialize a counter variable.
			Dim count As Integer = 1

			'Create a StringBuilder object to store the result.
			Dim builder As New StringBuilder()

			'Iterate through each TextSelection object in the array.
			For Each selection As TextSelection In textSelections
				'Get the layout entities (lines) associated with the first range in the TextSelection.
				For Each line As FixedLayoutSpan In layoutDoc.GetLayoutEntitiesOfNode(selection.GetRanges()(0))
					'Get the page index of the line.
					Dim index As Integer = line.PageIndex

				'Append the result to the StringBuilder.
				builder.AppendLine("The matched word " & count & " is on page:" & index)
				
				'Increment the counter.
				count += 1
			Next line
			Next selection

			'Save the result to a text file.
			File.WriteAllText("result.txt", builder.ToString())

			'Dispose of the Document object.
			document.Dispose()
			
			Process.Start("result.txt")
		End Sub

	End Class
End Namespace
