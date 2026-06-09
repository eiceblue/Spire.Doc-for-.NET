Imports System
Imports System.Windows.Forms
Imports Spire.Doc

Namespace HiddenRow
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new instance of the Document class
            Dim doc As Document = New Document()

            ' Load the content from the specified Word document file path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\TableTemplate.docx")

            ' Get the first section (index 0) from the document's sections collection
            Dim section As Section = doc.Sections[0]

            ' Cast the first table found in the section to a Table object
            Dim table As Table = (Table)section.Tables[0]

            ' Get the first row (index 0) from the table's rows collection
            Dim row As TableRow = table.Rows[0]

            ' Set the Hidden property to true to hide this row in the document
            row.Hidden = True

            ' Define the file path and name for the output document
            Dim result As String = "HiddenRow.docx"

            ' Save the modified document to a file in standard Docx format
            doc.SaveToFile(result, FileFormat.Docx)

            ' Close the document to release resources
            doc.Close()

            ' Dispose of the document object to free up memory
            doc.Dispose()

            'Launching the Word file.
            WordDocViewer(result)
        End Sub

        Private Sub WordDocViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

    End Class
End Namespace
