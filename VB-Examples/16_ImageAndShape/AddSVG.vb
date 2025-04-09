Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AddSVG
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input SVG file path
			Dim inputSvg As String = "../../../../../../Data/charthtml.svg"

			' Specify the output Word document file path
			Dim outputFile As String = "addSVG.docx"

			' Create a new Document object
			Dim document As New Document()

			' Add a new Section to the document
			Dim section As Section = document.AddSection()

			' Add a new Paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append the picture (SVG) to the paragraph
			paragraph.AppendPicture(inputSvg)

			' Save the document to the specified output file
			document.SaveToFile(outputFile, FileFormat.Docx2013)

			' Close the document
			document.Close()

			' Launch Word file
			WordDocViewer(outputFile)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace