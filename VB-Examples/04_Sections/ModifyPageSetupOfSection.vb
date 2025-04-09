Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ModifyPageSetupOfSection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object called doc
			Dim doc As New Document()

			'Load a Word document from the specified file path
			doc.LoadFromFile("../../../../../../Data/Template_N2.docx")

			'Iterate through each Section in the document
			For Each section As Section In doc.Sections

				'Set the margins of the current section using the MarginsF class
				section.PageSetup.Margins = New MarginsF(100, 80, 100, 80)

				'Set the page size of the current section to Letter
				section.PageSetup.PageSize = PageSize.Letter
			Next section

			' Or only modify one section
			' For example, modify the page setup of the first section
			'Section section0 = doc.Sections[0];
			'section0.PageSetup.Margins = new MarginsF(100, 80, 100, 80);
			'section0.PageSetup.FooterDistance = 35.4f;
			'section0.PageSetup.HeaderDistance = 34.4f;
			
			'Specify the output file name
			Dim output As String = "ModifyPageSetupOfAllSections_out.docx"

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
