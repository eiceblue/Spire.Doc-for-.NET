Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace PreserveTheme
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim input As String = "..\..\..\..\..\..\..\Data\Theme.docx"

			' Create a new Document object
			Dim doc As New Document()
			' Load the Word document from the specified input file path
			doc.LoadFromFile(input)

			' Create a new Document object for the modified document
			Dim newWord As New Document()
			' Clone the default style settings from the original document to the new document
			doc.CloneDefaultStyleTo(newWord)
			' Clone the themes from the original document to the new document
			doc.CloneThemesTo(newWord)
			' Clone the compatibility settings from the original document to the new document
			doc.CloneCompatibilityTo(newWord)

			' Add a clone of the first section from the original document to the new document
			newWord.Sections.Add(doc.Sections(0).Clone())

			' Specify the output file path
			Dim output As String = "PreserveTheme.docx"
			' Save the modified document to the specified output file path with the specified file format
			newWord.SaveToFile(output, FileFormat.Docx)

			' Dispose of the resources used by the Document objects
			doc.Dispose()
			newWord.Dispose()

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
