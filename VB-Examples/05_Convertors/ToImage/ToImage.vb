Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ConvertToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
            Dim document_Renamed As New Document()
            document_Renamed.LoadFromFile("..\..\..\..\..\..\Data\Word.doc")

			'Save doc file.
            Dim img As Image = document_Renamed.SaveToImages(0, ImageType.Metafile)
            img.Save("sample.bmp", System.Drawing.Imaging.ImageFormat.Bmp)

			'Launching the image file.
			WordDocViewer("sample.bmp")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
