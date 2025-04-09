Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SplitDocBySectionBreak
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\..\Data\Template_Docx_4.docx")

			' Declare a new Word document variable
			Dim newWord As Document

			' Iterate through each Section in the document
			For i As Integer = 0 To document.Sections.Count - 1
				' Set the file name for the result
				Dim result As String = String.Format("Result-SplitWordFileBySectionBreak_{0}.docx", i)
				
				' Create a new Document object for the split section
				newWord = New Document()
				
				' Add a cloned section from the original document to the new document
				newWord.Sections.Add(document.Sections(i).Clone())
				
				' Save the new document to a file
				newWord.SaveToFile(result)
				
				' Release all resources used by the new Document object
				newWord.Dispose()

				'Launch the MS Word file.
				WordDocViewer(result)
			Next i
			' Release all resources used by the original Document object
			document.Dispose()
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
