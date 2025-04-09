Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddAndDeleteSections
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object.
			Dim doc As New Document()

			' Load a Word document from a specified file path.
			doc.LoadFromFile("..\..\..\..\..\..\Data\SectionTemplate.docx")

			' Call the AddSection subroutine to add a new section to the document.
			AddSection(doc)

			' Call the DeleteSection subroutine to delete the last section of the document.
			DeleteSection(doc)

			' Specify the output file name for the modified document.
			Dim output As String = "AddAndDeleteSections_out.docx"

			' Save the modified document to the specified output file path in Docx2013 format.
			doc.SaveToFile(output, FileFormat.Docx2013)
			' Dispose the document object to release any resources it holds.
			doc.Dispose()

			FileViewer(output)
		End Sub
		' Subroutine to add a new section to the document.
		Private Sub AddSection(ByVal doc As Document)
			doc.AddSection()
		End Sub

		' Subroutine to delete the last section of the document.
		Private Sub DeleteSection(ByVal doc As Document)
			doc.Sections.RemoveAt(doc.Sections.Count - 1)
		End Sub
		
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
