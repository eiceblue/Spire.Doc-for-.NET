Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace CopyDocumentStyles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Load source document from disk
			Dim srcDoc As New Document()
			srcDoc.LoadFromFile("..\..\..\..\..\..\Data\Template_Toc.docx")

			'Load destination document from disk
			Dim destDoc As New Document()
			destDoc.LoadFromFile("..\..\..\..\..\..\Data\Template_N3.docx")

			'Get the style collections of source document
			Dim styles As Spire.Doc.Collections.StyleCollection = srcDoc.Styles

			'Add the style to destination document
			For Each style As Style In styles
				destDoc.Styles.Add(style)
			Next style

			'Save the Word file
			Dim output As String = "CopyDocumentStyles_out.docx"
			destDoc.SaveToFile(output, FileFormat.Docx2013)

			'Dispose of the document object
			srcDoc.Dispose()
			destDoc.Dispose()
			
			'Launch the file
			FileViewer(output)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
