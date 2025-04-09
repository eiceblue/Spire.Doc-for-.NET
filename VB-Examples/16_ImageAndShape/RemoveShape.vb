﻿Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace RemoveShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\Shapes.docx"

			'Create a word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first section
			Dim section As Section = doc.Sections(0)

			'Get all the child objects of paragraph
			For Each para As Paragraph In section.Paragraphs
				Dim i As Integer = 0
				Do While i < para.ChildObjects.Count

					'If the child objects is shape object
					If TypeOf para.ChildObjects(i) Is ShapeObject Then

						'Remove the shape object
						para.ChildObjects.RemoveAt(i)
					End If
					i += 1
				Loop
			Next para

			'Save and launch document
			Dim output As String = "RemoveShape.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
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
