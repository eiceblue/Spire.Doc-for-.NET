Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AddTableCaption
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim document As New Document()

			' Load the Word document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\Data\TableTemplate.docx")

			' Get the body of the first section in the document
			Dim body As Body = document.Sections(0).Body

			' Get the first table in the body
			Dim table As Table = TryCast(body.Tables(0), Table)

			' Add a caption to the table with the "Table" label, numbering format as "Number", and position below the table
			table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem)

			' Enable field updating in the document
			document.IsUpdateFields = True

			' Specify the output file name and format (Docx)
			Dim output As String = "AddTableCaption_result.docx"
			document.SaveToFile(output, FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launching the file
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
