Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.IO
Imports System.Text
Imports Spire.Doc.Fields

Namespace GetAlternativeText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a document
			Dim document As New Document()

			'Load the document from disk
			document.LoadFromFile("..\..\..\..\..\..\Data\ShapeWithAlternativeText.docx")

			'Create string builder
			Dim builder As New StringBuilder()

			'Loop through shapes and get the AlternativeText
			For Each section As Section In document.Sections

				'Loop through the paragraphs in the section
				For Each para As Paragraph In section.Paragraphs

					'//Loop through the child objects in the paragraph
					For Each obj As DocumentObject In para.ChildObjects

						'If the shape is a shape object
						If TypeOf obj Is ShapeObject Then
							Dim text As String = (TryCast(obj, ShapeObject)).AlternativeText
							'Append the alternative text in builder
							builder.AppendLine(text)
						End If
					Next obj
				Next para
			Next section

			'Write the content in txt file
			Dim result As String = "GetAlternativeText_result.txt"
			File.WriteAllText(result, builder.ToString())

			'Dispose the document
			document.Dispose()

			'Launch the file
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
