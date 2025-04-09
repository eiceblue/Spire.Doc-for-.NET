Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemovePageBreaks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load an existing document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_4.docx")

			' Iterate through all paragraphs in the first section of the document
			For j As Integer = 0 To document.Sections(0).Paragraphs.Count - 1
				Dim p As Paragraph = document.Sections(0).Paragraphs(j)

				' Iterate through all child objects within the paragraph
				Dim i As Integer = 0
				Do While i < p.ChildObjects.Count
					Dim obj As DocumentObject = p.ChildObjects(i)

					' Check if the child object is a page break and remove it if found
					If obj.DocumentObjectType = DocumentObjectType.Break Then
						Dim b As Break = TryCast(obj, Break)
						p.ChildObjects.Remove(b)
					End If
					i += 1
				Loop
			Next j

			' Specify the file name for the resulting document
			Dim result As String = "Result-RemovePageBreaks.docx"

			' Save the modified document to a new file with the specified file format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object
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
