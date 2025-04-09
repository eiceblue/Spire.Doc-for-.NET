Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace CloneSection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object called srcDoc
			Dim srcDoc As New Document()

			'Load a Word document from the specified file path
			srcDoc.LoadFromFile("..\..\..\..\..\..\Data\SectionTemplate.docx")

			'Create a new Document object called desDoc
			Dim desDoc As New Document()

			'Declare a Section object called cloneSection and initialize it as Nothing
			Dim cloneSection As Section = Nothing

			'Iterate through each Section in srcDoc
			For Each section As Section In srcDoc.Sections

				'Clone the current section and assign it to cloneSection
				cloneSection = section.Clone()

				'Add the cloned section to desDoc
				desDoc.Sections.Add(cloneSection)
			Next section

			'Specify the output file name
			Dim output As String = "CloneSection_out.docx"

			'Save the desDoc as a Word document with the specified file format
			desDoc.SaveToFile(output, FileFormat.Docx2013)

			'Dispose of the srcDoc object to release resources
			srcDoc.Dispose()

			'Dispose of the desDoc object to release resources
			desDoc.Dispose()

			'Launch Word file
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