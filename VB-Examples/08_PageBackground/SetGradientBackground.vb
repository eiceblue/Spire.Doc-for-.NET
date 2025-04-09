Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetGradientBackground
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the specified document file
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_2.docx")

			' Set the background type of the document to gradient
			document.Background.Type = BackgroundType.Gradient

			' Access the BackgroundGradient object from the document
			Dim Test As BackgroundGradient = document.Background.Gradient

			' Set the first color of the gradient background
			Test.Color1 = Color.White

			' Set the second color of the gradient background
			Test.Color2 = Color.LightBlue

			' Set the shading variant for the gradient background
			Test.ShadingVariant = GradientShadingVariant.ShadingDown

			' Set the shading style for the gradient background
			Test.ShadingStyle = GradientShadingStyle.Horizontal

			' Specify the file path for the output result
			Dim result As String = "Result-SetGradientBackground.docx"

			' Save the document to a file with the specified file format
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
