Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ReplaceWithHtml
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Read the contents of the input HTML file into a string.
			Dim HTML As String = File.ReadAllText("..\..\..\..\..\..\Data\InputHtml1.txt")

			' Create a new Document object and load a Word document to replace with HTML.
			Dim document As New Document("..\..\..\..\..\..\Data\ReplaceWithHtml.docx")

			' Create a list to store the replacement HTML content as DocumentObjects.
			Dim replacement As New List(Of DocumentObject)()

			' Create a temporary section in the document.
			Dim tempSection As Section = document.AddSection()

			' Add a paragraph to the temporary section and append the HTML content.
			Dim par As Paragraph = tempSection.AddParagraph()
			par.AppendHTML(HTML)

			' Retrieve the DocumentObjects from the temporary section and add them to the replacement list.
			For Each obj As DocumentObject In tempSection.Body.ChildObjects
				Dim docObj As DocumentObject = TryCast(obj, DocumentObject)
				replacement.Add(docObj)
			Next obj

			' Find all occurrences of the placeholder "[#placeholder]" in the document.
			Dim selections() As TextSelection = document.FindAllString("[#placeholder]", False, True)

			' Create a list to store the locations of the text range containing the placeholder.
			Dim locations As New List(Of TextRangeLocation)()
			For Each selection As TextSelection In selections
				locations.Add(New TextRangeLocation(selection.GetAsOneRange()))
			Next selection

			' Sort the locations based on the index of the text range.
			locations.Sort()

			' Iterate through each location and replace the text range with the HTML content.
			For Each location As TextRangeLocation In locations
				ReplaceWithHTML(location, replacement)
			Next location

			' Remove the temporary section from the document.
			document.Sections.Remove(tempSection)

			' Save the modified document to the specified output file in Docx format.
			document.SaveToFile("Output.docx", FileFormat.Docx)

			' Dispose of the Document object to release resources.
			document.Dispose()
			
			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub

		' Method to replace a text range with HTML content.
		Private Sub ReplaceWithHTML(ByVal location As TextRangeLocation, ByVal replacement As List(Of DocumentObject))
			Dim textRange As TextRange = location.Text
			Dim index As Integer = location.Index
			Dim paragraph As Paragraph = location.Owner
			Dim sectionBody As Body = paragraph.OwnerTextBody
			Dim paragraphIndex As Integer = sectionBody.ChildObjects.IndexOf(paragraph)
			Dim replacementIndex As Integer = -1

			If index = 0 Then
				' Remove the first child object (text range) from the paragraph.
				paragraph.ChildObjects.RemoveAt(0)
				replacementIndex = sectionBody.ChildObjects.IndexOf(paragraph)
			ElseIf index = paragraph.ChildObjects.Count - 1 Then
				' Remove the last child object (text range) from the paragraph.
				paragraph.ChildObjects.RemoveAt(index)
				replacementIndex = paragraphIndex + 1
			Else
				' Clone the current paragraph and split its child objects before and after the text range.
				Dim paragraph1 As Paragraph = CType(paragraph.Clone(), Paragraph)
				Do While paragraph.ChildObjects.Count > index
					paragraph.ChildObjects.RemoveAt(index)
				Loop
				Dim i As Integer = 0
				Dim count As Integer = index + 1
				Do While i < count
					paragraph1.ChildObjects.RemoveAt(0)
					i += 1
				Loop
				' Insert the cloned paragraph after the original paragraph within the body.
				sectionBody.ChildObjects.Insert(paragraphIndex + 1, paragraph1)
				replacementIndex = paragraphIndex + 1
			End If

			' Insert the replacement HTML content at the specified index within the body.
			For i As Integer = 0 To replacement.Count - 1
				sectionBody.ChildObjects.Insert(replacementIndex + i, replacement(i).Clone())
			Next i
		End Sub

		' Class representing the location of a text range within a paragraph.
		Public Class TextRangeLocation
			Implements IComparable(Of TextRangeLocation)

			Public Sub New(ByVal text As TextRange)
				Me.Text = text
			End Sub

			' Property for the TextRange object.
			Public Property Text() As TextRange
				Get
					Return m_Text
				End Get
				Set(ByVal value As TextRange)
					m_Text = value
				End Set
			End Property

			Private m_Text As TextRange

			' Property for retrieving the owner paragraph of the text range.
			Public ReadOnly Property Owner() As Paragraph
				Get
					Return Me.Text.OwnerParagraph
				End Get
			End Property

			' Property for retrieving the index of the text range within its owner paragraph.
			Public ReadOnly Property Index() As Integer
				Get
					Return Me.Owner.ChildObjects.IndexOf(Me.Text)
				End Get
			End Property

			' Implementation of the CompareTo method for sorting purposes.
			Public Function CompareTo(ByVal other As TextRangeLocation) As Integer Implements IComparable(Of TextRangeLocation).CompareTo
				' Compare two text range locations based on their indices (descending order).
				Return -(Me.Index - other.Index)
			End Function
		End Class

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
