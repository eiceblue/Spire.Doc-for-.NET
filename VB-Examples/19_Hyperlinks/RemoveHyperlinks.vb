Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace RemoveHyperlinks
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

		   ' Find all the hyperlinks in the document and store them in a list
		   Dim hyperlinks As List(Of Field) = FindAllHyperlinks(doc)

		   ' Flatten each hyperlink, removing the hyperlink functionality but keeping the text
		   For i As Integer = hyperlinks.Count - 1 To 0 Step -1
			   FlattenHyperlinks(hyperlinks(i))
		   Next i

		   ' Specify the output file path for the modified document without hyperlinks
		   Dim output As String = "RemoveHyperlinks.docx"

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

	   ' Method to find all hyperlinks in the document and return them as a list
	   Private Shared Function FindAllHyperlinks(ByVal document As Document) As List(Of Field)
		   Dim hyperlinks As New List(Of Field)()

		   For Each section As Section In document.Sections
			   For Each sec As DocumentObject In section.Body.ChildObjects
				   If sec.DocumentObjectType = DocumentObjectType.Paragraph Then
					   For Each para As DocumentObject In (TryCast(sec, Paragraph)).ChildObjects
						   If para.DocumentObjectType = DocumentObjectType.Field Then
							   Dim field As Field = TryCast(para, Field)
							   If field.Type = FieldType.FieldHyperlink Then
								   hyperlinks.Add(field)
							   End If
						   End If
					   Next para
				   End If
			   Next sec
		   Next section
		   Return hyperlinks
	   End Function


	   ' Method to flatten a hyperlink, removing the hyperlink functionality but keeping the text
	   Private Shared Sub FlattenHyperlinks(ByVal field As Field)
		   ' Store the indices of relevant objects for later removal
		   Dim ownerParaIndex As Integer = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.OwnerParagraph)
		   Dim fieldIndex As Integer = field.OwnerParagraph.ChildObjects.IndexOf(field)
		   Dim sepOwnerPara As Paragraph = field.Separator.OwnerParagraph
		   Dim sepOwnerParaIndex As Integer = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.Separator.OwnerParagraph)
		   Dim sepIndex As Integer = field.Separator.OwnerParagraph.ChildObjects.IndexOf(field.Separator)
		   Dim endIndex As Integer = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End)
		   Dim endOwnerParaIndex As Integer = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.End.OwnerParagraph)

		   ' Format the text between the separator and the end of the field result
		   FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex)

		   ' Remove the end field marker
		   field.End.OwnerParagraph.ChildObjects.RemoveAt(endIndex)

		   ' Remove the field and its associated objects in reverse order
		   For i As Integer = sepOwnerParaIndex To ownerParaIndex Step -1
			   If i = sepOwnerParaIndex AndAlso i = ownerParaIndex Then
				   ' Remove objects from the same paragraph as the field
				   For j As Integer = sepIndex To fieldIndex Step -1
					   field.OwnerParagraph.ChildObjects.RemoveAt(j)
				   Next j
			   ElseIf i = ownerParaIndex Then
				   ' Remove objects from the field's paragraph but after the field
				   For j As Integer = field.OwnerParagraph.ChildObjects.Count - 1 To fieldIndex Step -1
					   field.OwnerParagraph.ChildObjects.RemoveAt(j)
				   Next j
			   ElseIf i = sepOwnerParaIndex Then
				   ' Remove objects from the separator's paragraph
				   For j As Integer = sepIndex To 0 Step -1
					   sepOwnerPara.ChildObjects.RemoveAt(j)
				   Next j
			   Else
				   ' Remove objects from other paragraphs
				   field.OwnerParagraph.OwnerTextBody.ChildObjects.RemoveAt(i)
			   End If
		   Next i
	   End Sub

	   ' Method to format the text between the separator and the end of a field result in the document body
	   Private Shared Sub FormatFieldResultText(ByVal ownerBody As Body, ByVal sepOwnerParaIndex As Integer, ByVal endOwnerParaIndex As Integer, ByVal sepIndex As Integer, ByVal endIndex As Integer)
		   For i As Integer = sepOwnerParaIndex To endOwnerParaIndex
			   ' Get the paragraph at the current index
			   Dim para As Paragraph = TryCast(ownerBody.ChildObjects(i), Paragraph)

			   If i = sepOwnerParaIndex AndAlso i = endOwnerParaIndex Then
				   ' Format objects within the same paragraph as the separator and the end of the field
				   For j As Integer = sepIndex + 1 To endIndex - 1
					   FormatText(TryCast(para.ChildObjects(j), TextRange))
				   Next j
			   ElseIf i = sepOwnerParaIndex Then
				   ' Format objects after the separator in the separator's paragraph
				   For j As Integer = sepIndex + 1 To para.ChildObjects.Count - 1
					   FormatText(TryCast(para.ChildObjects(j), TextRange))
				   Next j
			   ElseIf i = endOwnerParaIndex Then
				   ' Format objects before the end of the field in the end paragraph
				   For j As Integer = 0 To endIndex - 1
					   FormatText(TryCast(para.ChildObjects(j), TextRange))
				   Next j
			   Else
				   ' Format all objects in other paragraphs
				   For j As Integer = 0 To para.ChildObjects.Count - 1
					   FormatText(TryCast(para.ChildObjects(j), TextRange))
				   Next j
			   End If
		   Next i
	   End Sub

	   ' Method to format the text range by setting its color to black and removing underline style
	   Private Shared Sub FormatText(ByVal tr As TextRange)
		   tr.CharacterFormat.TextColor = Color.Black
		   tr.CharacterFormat.UnderlineStyle = UnderlineStyle.None
	   End Sub
	End Class
End Namespace
