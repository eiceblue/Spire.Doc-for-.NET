Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertOLEAsIconViaStream
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the output file name
			Dim output As String = "InsertOLEAsIconViaStream.docx"

			' Create a new document object
			Dim doc As New Document()

			' Add a section to the document
			Dim sec As Section = doc.AddSection()

			' Add a paragraph to the section
			Dim par As Paragraph = sec.AddParagraph()

			' Open a stream for the OLE object data from the specified file
			Dim stream As Stream = File.OpenRead("..\..\..\..\..\..\Data\example.zip")

			' Create a DocPicture object and load an image from file
			Dim picture As New DocPicture(doc)
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\example.png")
			picture.LoadImage(image)

			' Append an OLE object to the paragraph using the provided stream, picture, and object type ("zip")
			Dim obj As DocOleObject = par.AppendOleObject(stream, picture, "zip")

			' Set the OLE object to be displayed as an icon
			obj.DisplayAsIcon = True

			' Save the document to a file in Docx2013 format
			doc.SaveToFile(output, FileFormat.Docx2013)

			' Dispose the document object
			doc.Dispose()

			'Launching the Word file.
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
