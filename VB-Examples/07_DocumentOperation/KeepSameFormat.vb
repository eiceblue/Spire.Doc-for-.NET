Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace KeepSameFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object for the source document
			Dim srcDoc As New Document()
			' Load the source document from the specified file path
			srcDoc.LoadFromFile("..\..\..\..\..\..\..\Data\Template_N2.docx")

			' Create a new Document object for the destination document
			Dim destDoc As New Document()
			' Load the destination document from the specified file path
			destDoc.LoadFromFile("..\..\..\..\..\..\..\Data\Template_N3.docx")

			' Set the KeepSameFormat property of the source document to true
			srcDoc.KeepSameFormat = True

			' Iterate through each Section in the source document
			For Each section As Section In srcDoc.Sections
				' Clone and add each Section to the destination document
				destDoc.Sections.Add(section.Clone())
			Next section

			' Specify the output file name and path
			Dim output As String = "KeepSameFormating_out.docx"
			' Save the destination document to the specified output file
			destDoc.SaveToFile(output, FileFormat.Docx2013)

			' Dispose the source document object
			srcDoc.Dispose()
			' Dispose the destination document object
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
