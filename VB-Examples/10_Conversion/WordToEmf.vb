Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.Drawing.Imaging

Namespace WordToEmf
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document.
			Dim document As New Document()

			'Load the file from disk.
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx", FileFormat.Docx)

			'Convert the first page of document to image.
			Dim image As Image = document.SaveToImages(0, Spire.Doc.Documents.ImageType.Metafile)
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim images As SkiaSharp.SKImage = document.SaveToImages(0, ImageType.Bitmap)
			'Using fileStream As New FileStream(outputFile, FileMode.Create, FileAccess.Write)
			'	images.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100).SaveTo(fileStream)
			'	fileStream.Flush()
			'End Using
			' =============================================================================

			' =============================================================================
			' Use the following code for WPF dlls
			' =============================================================================
			'Dim images() As BitmapSource = document.SaveToImages(Spire.Doc.Documents.ImageType.Bitmap)
			'Dim pE As New PngBitmapEncoder()
			'pE.Frames.Add(BitmapFrame.Create(images(0)))
			'Dim outputfile As String = String.Format(outputfile, ImageFormat.Png)
			'Using stream As Stream = File.Create(outputfile)
			'	pE.Save(stream)
			'End Using
			' =============================================================================

			Dim result As String = "Result-WordToEmf.emf"

			'Save the file.
			image.Save(result, ImageFormat.Emf)

			'Dispose the document.
			document.Dispose()

			'Launch the file.
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
