Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace HideParagraph
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			' Get the first Section in the document
			Dim sec As Section = document.Sections(0)

			' Get the first Paragraph in the section
			Dim para As Paragraph = sec.Paragraphs(0)

			' Iterate through each ChildObject in the paragraph
			For Each obj As DocumentObject In para.ChildObjects

				' Check if the current ChildObject is a TextRange
				If TypeOf obj Is TextRange Then
				
					' Cast the ChildObject to a TextRange
					Dim range As TextRange = TryCast(obj, TextRange)
					
					' Set the Hidden property of the TextRange to True (hide the text)
					range.CharacterFormat.Hidden = True
					
				End If
				
			Next obj

			' Specify the filename for the output result
			Dim result As String = "Result-HideWordParagraph.docx"

			' Save the modified document to the specified file in Docx2013 format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose the Document object to release resources
			document.Dispose()

			'Launch the MS Word file.
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
