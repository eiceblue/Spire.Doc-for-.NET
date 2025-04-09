Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace CloneSectionContent
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object called doc
			Dim doc As New Document()

			'Load a Word document from the specified file path
			doc.LoadFromFile("..\..\..\..\..\..\Data\SectionTemplate.docx")

			'Get the first section from the document and assign it to sec1
			Dim sec1 As Section = doc.Sections(0)

			'Get the second section from the document and assign it to sec2
			Dim sec2 As Section = doc.Sections(1)

			'Iterate through each DocumentObject in the body of sec1
			For Each obj As DocumentObject In sec1.Body.ChildObjects

				'Clone the current DocumentObject and add it to the body of sec2
				sec2.Body.ChildObjects.Add(obj.Clone())
			Next obj

			'Specify the output file name
			Dim output As String = "CloneSectionContent_out.docx"

			'Save the modified document as a Word document with the specified file format
			doc.SaveToFile(output, FileFormat.Docx2013)

			'Dispose of the doc object to release resources
			doc.Dispose()

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
