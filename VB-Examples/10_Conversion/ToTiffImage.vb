Imports System.ComponentModel
Imports System.Drawing.Imaging
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ToTiffImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
			Dim document As New Document()

			'Load the file from disk
			document.LoadFromFile("..\..\..\..\..\..\Data\ConvertedTemplate.docx")

			'Save the document to a tiff image.
			JoinTiffImages(SaveAsImage(document),"Sample.tif",EncoderValue.CompressionLZW)

			'Dispose the document
			document.Dispose()

			'Launching the tiff file.
			FileViewer("Sample.tif")
		End Sub

		Private Shared Function SaveAsImage(ByVal document As Document) As Image()
			'Save all the pages in the document to images.
			Dim images() As Image = document.SaveToImages(ImageType.Bitmap)
			Return images
		End Function

		Private Shared Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
			'Set the code information for the images.
			Dim encoders() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders()
			For j As Integer = 0 To encoders.Length - 1
				If encoders(j).MimeType = mimeType Then
					Return encoders(j)
				End If
			Next j
			Throw New Exception(mimeType & " mime type not found in ImageCodecInfo")
		End Function

		Public Shared Sub JoinTiffImages(ByVal images() As Image, ByVal outFile As String, ByVal compressEncoder As EncoderValue)
			'Set the encoder parameters.
			Dim enc As System.Drawing.Imaging.Encoder = System.Drawing.Imaging.Encoder.SaveFlag
			Dim ep As New EncoderParameters(2)
			ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.MultiFrame))
			ep.Param(1) = New EncoderParameter(System.Drawing.Imaging.Encoder.Compression, CLng(compressEncoder))
			Dim pages As Image = images(0)
			Dim frame As Integer = 0
			Dim info As ImageCodecInfo = GetEncoderInfo("image/tiff")
			For Each img As Image In images
				If frame = 0 Then
					pages = img
					'Save the first frame.
					pages.Save(outFile, info, ep)

				Else
					'Save the intermediate frames.
					ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.FrameDimensionPage))

					pages.SaveAdd(img, ep)
				End If
				If frame = images.Length - 1 Then
					'Flush and close.
					ep.Param(0) = New EncoderParameter(enc, CLng(EncoderValue.Flush))
					pages.SaveAdd(ep)
				End If
				frame += 1
			Next img
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
