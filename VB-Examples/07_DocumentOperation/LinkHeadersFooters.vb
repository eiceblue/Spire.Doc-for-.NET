Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace LinkHeadersFooters
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new instance of the Document class and assign it to the srcDoc variable.
			Dim srcDoc As New Document()

			'Load a Word document from the specified file path and assign it to the srcDoc variable.
			srcDoc.LoadFromFile("..\..\..\..\..\..\..\Data\Template_N1.docx")

			'Create a new instance of the Document class and assign it to the dstDoc variable.
			Dim dstDoc As New Document()

			'Load a Word document from the specified file path and assign it to the dstDoc variable.
			dstDoc.LoadFromFile("..\..\..\..\..\..\..\Data\Template_N2.docx")

			'Set the LinkToPrevious property of the header in the first section of srcDoc to True.
			srcDoc.Sections(0).HeadersFooters.Header.LinkToPrevious = True

			'Set the LinkToPrevious property of the footer in the first section of srcDoc to True.
			srcDoc.Sections(0).HeadersFooters.Footer.LinkToPrevious = True

			'Iterate through each section in srcDoc and add a cloned section to dstDoc.
			For Each section As Section In srcDoc.Sections
				dstDoc.Sections.Add(section.Clone())
			Next section

			'Specify the output file name as "LinkHeadersFooters_out.docx".
			Dim output As String = "LinkHeadersFooters_out.docx"

			'Save the dstDoc document to the specified file path in the Docx2013 file format.
			dstDoc.SaveToFile(output, FileFormat.Docx2013)

			'Dispose of the resources used by srcDoc.
			srcDoc.Dispose()

			'Dispose of the resources used by dstDoc.
			dstDoc.Dispose()

			'Launching the document
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
