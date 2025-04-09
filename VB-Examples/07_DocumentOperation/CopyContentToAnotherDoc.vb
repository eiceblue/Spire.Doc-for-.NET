Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace CopyContentToAnotherDoc
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Initialize a new instance of the Document class.
			Dim sourceDoc As New Document()

			'Load the source document from the specified file path.
			sourceDoc.LoadFromFile("..\..\..\..\..\..\..\Data\Template_Docx_1.docx")

			'Initialize a new instance of the Document class.
			Dim destinationDoc As New Document()

			'Load the destination document from the specified file path.
			destinationDoc.LoadFromFile("..\..\..\..\..\..\..\Data\Target.docx")

			'Iterate through each Section in the source document.
			For Each sec As Section In sourceDoc.Sections

			'Iterate through each DocumentObject in the Body of the section.
			For Each obj As DocumentObject In sec.Body.ChildObjects

				'Clone the DocumentObject and add it to the Body of the first Section in the destination document.
				destinationDoc.Sections(0).Body.ChildObjects.Add(obj.Clone())

			Next obj
			Next sec

			'Specify the file name for the resulting document.
			Dim result As String = "Result-CopyContentToAnotherWord.docx"

			'Save the destination document to the specified file path in Docx2013 format.
			destinationDoc.SaveToFile(result, FileFormat.Docx2013)

			'Dispose of the source document object.
			sourceDoc.Dispose()

			'Dispose of the destination document object.
			destinationDoc.Dispose()

			'Launch the MS Word file.
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
