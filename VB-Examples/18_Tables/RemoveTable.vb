Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemoveTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            ' Define the input file path
            Dim input As String = "..\..\..\..\..\..\..\Data\Template.docx"

            ' Create a new Document object
            Dim doc As New Document()

            ' Load an existing Word document from a file
            doc.LoadFromFile(input)

            'Remove the first Table            
            doc.Sections(0).Tables.RemoveAt(0)

            'Save the document
            Dim output As String = "RemoveTable.docx"
            doc.SaveToFile(output, FileFormat.Docx)

            'Dispose of the document object to free up resources
            doc.Dispose()
            Viewer(output)
        End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
