Imports Spire.Doc

Namespace SetGutterPosition
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			' Access the first section of the document
			Dim section As Section = document.Sections(0)

			' Set the IsTopGutter property to True for the page setup of the section
			section.PageSetup.IsTopGutter = True

			' Set the Gutter width to 100.0F for the page setup of the section
			section.PageSetup.Gutter = 100.0F

			' Define the output file name for the modified document
			Dim output As String = "SetGutterPosition.docx"

			' Save the modified document to the specified output file path in Docx format
			document.SaveToFile(output, FileFormat.Docx)

			' Dispose of the Document object to release resources
			document.Dispose()

			'Launch the file
			WordDocViewer(output)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
