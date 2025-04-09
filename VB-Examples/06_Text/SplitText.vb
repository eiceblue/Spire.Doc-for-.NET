Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SplitText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			 ' Set the input file path for the Word document
			Dim input As String = "..\..\..\..\..\..\Data\Sample.docx"

			' Create a new instance of the Document class
			Dim doc As New Document()

			' Load the Word document from the specified input file
			doc.LoadFromFile(input)

			' Add a column to the first section of the document with width 100.0F and spacing 20.0F
			doc.Sections(0).AddColumn(100.0F, 20.0F)

			' Enable the display of a line between columns in the first section's page setup
			doc.Sections(0).PageSetup.ColumnsLineBetween = True

			' Set the output file path for the modified Word document
			Dim output As String = "SplitText.docx"

			' Save the modified document to the specified output file in DOCX format compatible with Word 2013
			doc.SaveToFile(output, FileFormat.Docx2013)

			' Dispose of the document object to release its resources
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
