Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace IntegrateFontTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class and load a document from the specified file path ("Template_N2.docx").
			Dim destDoc As New Document()
			destDoc.LoadFromFile("..\..\..\..\..\..\..\Data\Template_N3.docx")


			' Create a new instance of the Document class and load another document from the specified file path ("Template_N3.docx").
			Dim srcDoc As New Document()
			srcDoc.LoadFromFile("..\..\..\..\..\..\..\Data\Template_N2.docx")


			' Integrate the current document font table to the destination document
			srcDoc.IntegrateFontTableTo(destDoc)

			' Iterate through each section in the source document.
			For Each section As Section In srcDoc.Sections
				' Clone each section and add it to the destination document.
				destDoc.Sections.Add(section.Clone())
			Next section

			' Specify the output file name.
			Dim output As String = "integrateFontTable.docx"

			' Save the modified destination document to a file with the specified output file name and format (Docx2013).
			destDoc.SaveToFile(output, FileFormat.Docx2016)

			' Clean up resources used by the source document.
			srcDoc.Dispose()

			' Clean up resources used by the destination document.
			destDoc.Dispose()

			'Launch the file
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
