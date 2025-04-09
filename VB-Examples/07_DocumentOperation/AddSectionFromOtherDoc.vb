Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.Text

Namespace AddSectionFromOtherDoc
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object and load a document from the specified file path.
			Dim TarDoc As New Document("..\..\..\..\..\..\..\Data\SampleB_1.docx")

			' Create another Document object and load a different document from the specified file path.
			Dim SouDoc As New Document("..\..\..\..\..\..\..\Data\Sample_two sections.docx")

			' Get the second section from the source document.
			Dim Ssection As Section = SouDoc.Sections(1)

			' Clone the second section and add it to the target document.
			TarDoc.Sections.Add(Ssection.Clone())

			' Specify the file name for the resulting document.
			Dim result As String = "result.docx"

			' Save the target document to the specified file path in Docx format.
			TarDoc.SaveToFile(result, FileFormat.Docx)

			' Dispose of the target document object to free up resources.
			TarDoc.Dispose()

			' Dispose of the source document object to free up resources.
			SouDoc.Dispose()

			'Launch result file
			WordDocViewer(result)

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
