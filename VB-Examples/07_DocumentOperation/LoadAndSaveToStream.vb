Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO
Imports Spire.Doc.Fields
Namespace LoadAndSaveToStream
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Set the input file path
			Dim input As String = "..\..\..\..\..\..\Data\Template.docx"

			' Open a stream to read the input file
			Dim stream As Stream = File.OpenRead(input)

			' Create a new Document object using the stream
			Dim doc As New Document(stream)

			' Close the stream
			stream.Close()

			' Do something with the document

			' Create a new memory stream to store the document content
			Dim newStream As New MemoryStream()

			' Save the document content to the new memory stream in RTF format
			doc.SaveToStream(newStream, FileFormat.Rtf)

			' Reset the position of the new memory stream to the beginning
			newStream.Position = 0

			' Set the file name for the result
			Dim result As String = "LoadAndSaveToStream_out.rtf"

			' Write the content of the new memory stream to a new file
			File.WriteAllBytes(result, newStream.ToArray())

			' Release all resources used by the Document object
			doc.Dispose()

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
