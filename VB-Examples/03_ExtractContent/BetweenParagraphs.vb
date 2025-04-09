Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace BetweenParagraphs
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object for the source document.
			Dim sourceDocument As New Document()

			' Load the source Word document from the specified file path.
			sourceDocument.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			' Create a new Document object for the destination document.
			Dim destinationDoc As New Document()

			' Create a new Section within the destination document.
			Dim section As Section = destinationDoc.AddSection()

			' Extract the content between specified paragraphs from the source document and add it to the destination document.
			ExtractBetweenParagraphs(sourceDocument, destinationDoc, 1, 3)

			' Save the modified destination document to the specified file in Docx format.
			destinationDoc.SaveToFile("Output.docx", FileFormat.Docx)

			' Dispose of the source and destination Document objects to release resources.
			sourceDocument.Dispose()
			destinationDoc.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub
		
		' Method to extract content between specified paragraphs from the source document and add it to the destination document.
		Private Sub ExtractBetweenParagraphs(ByVal sourceDocument As Document, ByVal destinationDocument As Document, ByVal startPara As Integer, ByVal endPara As Integer)
			For i As Integer = startPara - 1 To endPara - 1
				' Clone the child object (paragraph or other) at the specified index in the source document.
				Dim docObj As DocumentObject = sourceDocument.Sections(0).Body.ChildObjects(i).Clone()

				' Add the cloned child object to the body of the first section in the destination document.
				destinationDocument.Sections(0).Body.ChildObjects.Add(docObj)
			Next i
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
