Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO

Namespace ConvertDocToByte
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path for the original document
			Dim input As String = "..\..\..\..\..\..\Data\Template.docx"

			' Create a new Document object
			Dim doc As New Document()

			' Load the original document from the specified input file path
			doc.LoadFromFile(input)

			' Create a new memory stream to save the document in Docx format
			Dim outStream As New MemoryStream()
			doc.SaveToStream(outStream, FileFormat.Docx)

			' Convert the contents of the memory stream to a byte array
			Dim docBytes() As Byte = outStream.ToArray()

			' The bytes are now ready to be stored/transmitted.

			' Create a new memory stream using the byte array of the original document
			Dim inStream As New MemoryStream(docBytes)

			' Create a new Document object using the memory stream containing the original document's byte array
			Dim newDoc As New Document(inStream)

			' Dispose of the original and new Document objects to release resources
			doc.Dispose()
			newDoc.Dispose()

		End Sub
	End Class
End Namespace
