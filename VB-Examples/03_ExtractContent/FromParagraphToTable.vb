Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace FromParagraphToTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Creating a new instance of the Document class for the source document.
			Dim sourceDocument As New Document()
			' Loading the source document from a file located at a specified path.
			sourceDocument.LoadFromFile("..\..\..\..\..\..\Data\IncludingTable.docx")

			' Creating a new instance of the Document class for the destination document.
			Dim destinationDoc As New Document()

			' Adding a new section to the destination document.
			Dim destinationSection As Section = destinationDoc.AddSection()

			' Calling the ExtractByTable method to extract content from the source document and add it to the destination document.
			ExtractByTable(sourceDocument, destinationDoc, 1, 1)

			' Saving the destination document to a file named "Output.docx" in the Docx file format.
			destinationDoc.SaveToFile("Output.docx", FileFormat.Docx)

			' Disposing of the resources used by the source document.
			sourceDocument.Dispose()
			' Disposing of the resources used by the destination document.
			destinationDoc.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub

		Private Sub ExtractByTable(ByVal sourceDocument As Document, ByVal destinationDocument As Document, ByVal startPara As Integer, ByVal tableNo As Integer)
			' Retrieving the table at the specified index from the first section of the source document.
			Dim table As Table = TryCast(sourceDocument.Sections(0).Tables(tableNo - 1), Table)

			' Getting the index of the table in the child objects collection of the body of the first section.
			Dim index As Integer = sourceDocument.Sections(0).Body.ChildObjects.IndexOf(table)

			' Iterating through the child objects of the body from the start paragraph to the index of the table.
			For i As Integer = startPara - 1 To index
				' Cloning the document object at the current index.
				Dim dobj As DocumentObject = sourceDocument.Sections(0).Body.ChildObjects(i).Clone()

				' Adding the cloned document object to the child objects collection of the body of the first section in the destination document.
				destinationDocument.Sections(0).Body.ChildObjects.Add(dobj)
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
