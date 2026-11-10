Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts

Namespace MoveToDocument
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new empty document instance.
			Dim doc As New Document()

			' Create a document navigator to help navigate and modify the document content.
			Dim navigator As New DocumentNavigator(doc)

			' Load an existing Word document from the specified relative file path.
			doc.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			' Move the cursor to the very beginning of the document.
			navigator.MoveToDocumentStart()

			' Write a new line of text at the start of the document.
			navigator.Writeln("Insert the content at the beginning of the document.")

			' Write another line of text immediately after the previous one at the start.
			navigator.Writeln("This is new content.")

			' Move the cursor to the very end of the document.
			navigator.MoveToDocumentEnd()

			' Insert a blank line at the end of the document.
			navigator.Writeln()

			' Insert a new line of text at the end of the document.
			navigator.Writeln("Insert the content at the end of the document.")

			' Save the modified document to a new file.
			doc.SaveToFile("MoveToDocument.docx", FileFormat.Docx)

			' Close the document to release internal resources.
			doc.Close()

			' Explicitly dispose of the document object to free memory immediately.
			doc.Dispose()

			'Launch the MS Word file.
			WordDocViewer("MoveToDocument.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
